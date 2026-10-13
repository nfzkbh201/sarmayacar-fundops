"""Run behind HTTPS with a persistent, private data directory."""
from __future__ import annotations

import hashlib
import hmac
import json
import os
from pathlib import Path
import re
import secrets
import sqlite3
import time
from contextlib import contextmanager

from fastapi import FastAPI, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware

SCHEMA = "sarmayacar.quarterly_report.v1"
MAX_BODY = 5 * 1024 * 1024
RETENTION_SECONDS = 24 * 60 * 60
LEASE_SECONDS = 120


def digest(value: str) -> str:
    return hashlib.sha256(value.encode()).hexdigest()


class Store:
    def __init__(self, path: Path):
        path.parent.mkdir(parents=True, exist_ok=True, mode=0o700)
        self.path = path
        with self.db() as db:
            db.executescript("""
                CREATE TABLE IF NOT EXISTS connections (
                    id TEXT PRIMARY KEY, name TEXT NOT NULL, token_hash TEXT NOT NULL,
                    target TEXT, last_seen REAL, revoked INTEGER NOT NULL DEFAULT 0
                );
                CREATE TABLE IF NOT EXISTS jobs (
                    id TEXT PRIMARY KEY, connection_id TEXT NOT NULL, request_id TEXT NOT NULL,
                    quarter TEXT NOT NULL, payload TEXT, payload_hash TEXT NOT NULL,
                    status TEXT NOT NULL, created REAL NOT NULL, updated REAL NOT NULL,
                    runner TEXT, lease_until REAL, result TEXT,
                    UNIQUE(connection_id, request_id)
                );
            """)
        path.chmod(0o600)

    @contextmanager
    def db(self):
        db = sqlite3.connect(self.path, timeout=10)
        db.row_factory = sqlite3.Row
        try:
            with db:
                yield db
        finally:
            db.close()

    def expire(self, db):
        now = time.time()
        db.execute("""UPDATE jobs SET payload=NULL, status='expired', updated=?
            WHERE created < ? AND status IN ('queued', 'processing')""",
            (now, now - RETENTION_SECONDS))
        db.execute("UPDATE jobs SET payload=NULL WHERE created < ?", (now - RETENTION_SECONDS,))

    @staticmethod
    def job_info(row):
        if not row:
            return None
        result = dict(row)
        for key in ("payload", "payload_hash", "runner", "lease_until", "request_id"):
            result.pop(key, None)
        result["result"] = json.loads(result["result"]) if result["result"] else None
        return result


def create_app(database: Path | None = None, publisher_token: str | None = None):
    token = publisher_token or os.environ.get("FIGMA_BRIDGE_PUBLISHER_TOKEN", "")
    if len(token) < 32:
        raise RuntimeError("FIGMA_BRIDGE_PUBLISHER_TOKEN must contain at least 32 characters.")
    store = Store(database or Path(os.environ.get("FIGMA_BRIDGE_DB", ".figma_bridge/reports.sqlite3")))
    app = FastAPI(docs_url=None, redoc_url=None, openapi_url=None)
    # Figma's UI has an opaque origin. Authentication uses headers, never cookies.
    app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_credentials=False,
        allow_methods=["GET", "POST", "DELETE"],
        allow_headers=["Authorization", "Content-Type"])

    @app.middleware("http")
    async def private_responses(request, call_next):
        response = await call_next(request)
        response.headers["Cache-Control"] = "no-store"
        response.headers["X-Content-Type-Options"] = "nosniff"
        return response

    def auth(request, publisher=False):
        supplied = request.headers.get("authorization", "").removeprefix("Bearer ")
        if publisher:
            if not hmac.compare_digest(supplied.encode(), token.encode()):
                raise HTTPException(401, "Invalid publisher credentials.")
            return None
        connection_id, separator, reader_token = supplied.partition(".")
        with store.db() as db:
            row = db.execute("SELECT * FROM connections WHERE id=? AND revoked=0", (connection_id,)).fetchone()
        if not separator or not row or not hmac.compare_digest(row["token_hash"], digest(reader_token)):
            raise HTTPException(401, "Connection expired or revoked. Pair again in Streamlit.")
        return dict(row)

    async def body(request):
        raw = bytearray()
        async for chunk in request.stream():
            raw.extend(chunk)
            if len(raw) > MAX_BODY:
                raise HTTPException(413, "Report exceeds the 5 MB transfer limit.")
        try:
            def reject_constant(value):
                raise ValueError("Non-finite numbers are not supported.")
            value = json.loads(raw, parse_constant=reject_constant)
        except (ValueError, UnicodeDecodeError):
            raise HTTPException(400, "Invalid JSON.")
        if not isinstance(value, dict):
            raise HTTPException(400, "Expected a JSON object.")
        return value

    def string(data, key, limit=200):
        value = data.get(key)
        if not isinstance(value, str) or not value.strip() or len(value) > limit:
            raise HTTPException(400, f"Invalid {key}.")
        return value.strip()

    @app.get("/health")
    def health():
        return {"status": "ok"}

    @app.post("/v1/connections")
    async def create_connection(request: Request):
        auth(request, publisher=True)
        name = string(await body(request), "name", 100)
        connection_id, reader_token = secrets.token_urlsafe(18), secrets.token_urlsafe(32)
        with store.db() as db:
            db.execute("INSERT INTO connections(id,name,token_hash) VALUES(?,?,?)",
                       (connection_id, name, digest(reader_token)))
        return {"id": connection_id, "name": name, "pairing_code": f"{connection_id}.{reader_token}"}

    @app.get("/v1/connections")
    def connections(request: Request):
        auth(request, publisher=True)
        with store.db() as db:
            store.expire(db)
            rows = db.execute("SELECT id,name,target,last_seen FROM connections WHERE revoked=0 ORDER BY name").fetchall()
            result = []
            for row in rows:
                item = dict(row)
                item["target"] = json.loads(item["target"]) if item["target"] else None
                latest = db.execute("SELECT * FROM jobs WHERE connection_id=? ORDER BY created DESC LIMIT 1", (row["id"],)).fetchone()
                item["latest_job"] = store.job_info(latest)
                item["online"] = bool(row["last_seen"] and row["last_seen"] > time.time() - 45)
                result.append(item)
        return {"connections": result}

    @app.delete("/v1/connections/{connection_id}")
    def revoke(connection_id: str, request: Request):
        auth(request, publisher=True)
        with store.db() as db:
            db.execute("UPDATE connections SET revoked=1 WHERE id=?", (connection_id,))
            db.execute("UPDATE jobs SET status='cancelled',payload=NULL WHERE connection_id=? AND status IN ('queued','processing')", (connection_id,))
        return {"status": "revoked"}

    @app.post("/v1/connections/{connection_id}/jobs")
    async def publish(connection_id: str, request: Request):
        auth(request, publisher=True)
        data = await body(request)
        request_id = string(data, "request_id", 100)
        payload = data.get("payload")
        if not isinstance(payload, dict) or payload.get("schema_version") != SCHEMA:
            raise HTTPException(400, "Unsupported report schema.")
        report = payload.get("report")
        quarter = report.get("quarter") if isinstance(report, dict) else None
        if not isinstance(quarter, str) or not re.fullmatch(r"Q[1-4] 20\d{2}", quarter):
            raise HTTPException(400, "Report must specify its quarter and year.")
        months = report.get("months")
        if not isinstance(months, list) or not months or not all(isinstance(m, str) for m in months):
            raise HTTPException(400, "Report months are missing.")
        if not any(payload.get(k) for k in ("companies", "operator_angels", "financial_exhibits", "portfolio_summary")):
            raise HTTPException(400, "Report contains no pages.")
        try:
            encoded = json.dumps(payload, sort_keys=True, separators=(",", ":"), allow_nan=False)
        except (TypeError, ValueError):
            raise HTTPException(400, "Report values must be valid JSON numbers and text.")
        now, job_id = time.time(), secrets.token_urlsafe(18)
        with store.db() as db:
            db.execute("BEGIN IMMEDIATE")
            store.expire(db)
            connection = db.execute("SELECT * FROM connections WHERE id=? AND revoked=0", (connection_id,)).fetchone()
            if not connection or not connection["target"]:
                raise HTTPException(409, "Pair the Figma plugin before sending a report.")
            existing = db.execute("SELECT * FROM jobs WHERE connection_id=? AND request_id=?", (connection_id, request_id)).fetchone()
            if existing:
                if existing["payload_hash"] != digest(encoded):
                    raise HTTPException(409, "This request was already used for a different report.")
                return store.job_info(existing)
            active = db.execute("SELECT id FROM jobs WHERE connection_id=? AND status IN ('queued','processing')", (connection_id,)).fetchone()
            if active:
                raise HTTPException(409, "A report is already waiting or importing. Wait for it to finish.")
            db.execute("INSERT INTO jobs(id,connection_id,request_id,quarter,payload,payload_hash,status,created,updated) VALUES(?,?,?,?,?,?,'queued',?,?)",
                       (job_id, connection_id, request_id, quarter, encoded, digest(encoded), now, now))
            return store.job_info(db.execute("SELECT * FROM jobs WHERE id=?", (job_id,)).fetchone())

    @app.post("/v1/plugin/connect")
    async def connect(request: Request):
        connection = auth(request)
        data = await body(request)
        target = {key: string(data, key) for key in ("target_id", "page_id", "page_name", "file_name")}
        with store.db() as db:
            db.execute("BEGIN IMMEDIATE")
            row = db.execute("SELECT target FROM connections WHERE id=?", (connection["id"],)).fetchone()
            current = json.loads(row["target"]) if row["target"] else None
            if current and (current["target_id"], current["page_id"]) != (target["target_id"], target["page_id"]):
                raise HTTPException(409, "This connection belongs to another Figma page. Create a new connection.")
            db.execute("UPDATE connections SET target=?,last_seen=? WHERE id=?", (json.dumps(target), time.time(), connection["id"]))
        return {"name": connection["name"], "target": target}

    @app.get("/v1/plugin/jobs")
    def pending(request: Request):
        connection = auth(request)
        with store.db() as db:
            store.expire(db)
            db.execute("UPDATE connections SET last_seen=? WHERE id=?", (time.time(), connection["id"]))
            row = db.execute("SELECT * FROM jobs WHERE connection_id=? AND status IN ('queued','processing') ORDER BY created LIMIT 1", (connection["id"],)).fetchone()
        return {"job": store.job_info(row)}

    @app.post("/v1/plugin/jobs/{job_id}/claim")
    async def claim(job_id: str, request: Request):
        connection = auth(request)
        data = await body(request)
        runner = string(data, "runner", 100)
        now = time.time()
        with store.db() as db:
            db.execute("BEGIN IMMEDIATE")
            store.expire(db)
            row = db.execute("SELECT * FROM jobs WHERE id=? AND connection_id=?", (job_id, connection["id"])).fetchone()
            if not row or row["status"] not in ("queued", "processing"):
                raise HTTPException(409, "Report is no longer waiting for import.")
            if row["runner"] and row["runner"] != runner and row["lease_until"] > now:
                raise HTTPException(409, "Another plugin session is importing this report.")
            db.execute("UPDATE jobs SET status='processing',runner=?,lease_until=?,updated=? WHERE id=?", (runner, now + LEASE_SECONDS, now, job_id))
            db.execute("UPDATE connections SET last_seen=? WHERE id=?", (now, connection["id"]))
            return {"id": job_id, "payload": json.loads(row["payload"])}

    @app.post("/v1/plugin/jobs/{job_id}/result")
    async def complete(job_id: str, request: Request):
        connection = auth(request)
        data = await body(request)
        runner = string(data, "runner", 100)
        status = data.get("status")
        if status not in ("succeeded", "failed"):
            raise HTTPException(400, "Invalid result status.")
        count = data.get("page_count", 0)
        if type(count) is not int or not 0 <= count <= 1000:
            raise HTTPException(400, "Invalid page count.")
        result = {"page_count": count, "message": str(data.get("message", ""))[:1000],
                  "warning_count": len(data.get("warnings", [])) if isinstance(data.get("warnings"), list) else 0}
        with store.db() as db:
            db.execute("BEGIN IMMEDIATE")
            row = db.execute("SELECT * FROM jobs WHERE id=? AND connection_id=?", (job_id, connection["id"])).fetchone()
            if not row or row["runner"] != runner:
                raise HTTPException(409, "Import session no longer owns this report.")
            if row["status"] == status:
                return store.job_info(row)
            if row["status"] != "processing":
                raise HTTPException(409, "Report is no longer importing.")
            db.execute("UPDATE jobs SET status=?,result=?,payload=NULL,updated=? WHERE id=?",
                       (status, json.dumps(result), time.time(), job_id))
            return store.job_info(db.execute("SELECT * FROM jobs WHERE id=?", (job_id,)).fetchone())

    return app
