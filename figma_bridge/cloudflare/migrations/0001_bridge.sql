CREATE TABLE connections (id TEXT PRIMARY KEY, name TEXT NOT NULL, token_hash TEXT NOT NULL, target TEXT, last_seen REAL, revoked INTEGER NOT NULL DEFAULT 0);
CREATE TABLE jobs (id TEXT PRIMARY KEY, connection_id TEXT NOT NULL REFERENCES connections(id), request_id TEXT NOT NULL, quarter TEXT NOT NULL, payload_hash TEXT NOT NULL, status TEXT NOT NULL CHECK (status IN ('queued', 'processing', 'succeeded', 'failed', 'expired', 'cancelled')), created REAL NOT NULL, updated REAL NOT NULL, runner TEXT, lease_until REAL, result TEXT, UNIQUE(connection_id, request_id));
CREATE UNIQUE INDEX one_active_job ON jobs(connection_id) WHERE status IN ('queued', 'processing');
CREATE INDEX latest_job ON jobs(connection_id, created DESC);
CREATE INDEX expiring_job ON jobs(created) WHERE status IN ('queued', 'processing');
CREATE TABLE payload_chunks (job_id TEXT NOT NULL REFERENCES jobs(id) ON DELETE CASCADE, part INTEGER NOT NULL, content TEXT NOT NULL, PRIMARY KEY(job_id, part));
CREATE TRIGGER clear_completed_payload AFTER UPDATE OF status ON jobs WHEN NEW.status NOT IN ('queued', 'processing') BEGIN DELETE FROM payload_chunks WHERE job_id = NEW.id; END;
