"""Start an isolated local bridge and handoff screen, without publishing reports."""
import argparse
import json
import os
from pathlib import Path
import secrets
import socket
import subprocess
import sys
import time
from urllib.request import urlopen

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / ".figma_bridge"


def free_port(preferred):
    with socket.socket() as sock:
        try:
            sock.bind(("127.0.0.1", preferred))
        except OSError:
            sock.bind(("127.0.0.1", 0))
        return sock.getsockname()[1]


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--pack", type=Path)
    args = parser.parse_args()
    if args.pack and not args.pack.is_file():
        parser.error("The data pack file does not exist.")
    DATA.mkdir(mode=0o700, exist_ok=True)
    config_path = DATA / "local.json"
    if config_path.exists():
        config = json.loads(config_path.read_text())
    else:
        config = {"publisher_token": secrets.token_urlsafe(40), "access_code": secrets.token_urlsafe(18)}
    service_port, ui_port = free_port(8765), free_port(8503)
    config.update(service_url=f"http://127.0.0.1:{service_port}", ui_url=f"http://127.0.0.1:{ui_port}")
    env = {**os.environ, "FIGMA_BRIDGE_URL": config["service_url"],
        "FIGMA_BRIDGE_PUBLISHER_TOKEN": config["publisher_token"], "FIGMA_BRIDGE_ACCESS_CODE": config["access_code"],
        "FIGMA_BRIDGE_DB": str(DATA / "reports.sqlite3")}
    if args.pack:
        env["FIGMA_PREVIEW_PACK"] = str(args.pack.resolve())
    commands = [
        [sys.executable, "-m", "uvicorn", "figma_bridge.service:create_app", "--factory", "--host", "127.0.0.1", "--port", str(service_port), "--no-access-log"],
        [sys.executable, "-m", "streamlit", "run", "tools/figma_connection_preview.py", "--server.address", "127.0.0.1", "--server.port", str(ui_port), "--server.headless", "true", "--browser.gatherUsageStats", "false"],
    ]
    processes = []
    try:
        for name, command in zip(("bridge", "preview"), commands):
            with (DATA / f"{name}.log").open("ab") as log:
                processes.append(subprocess.Popen(command, cwd=ROOT, env=env, stdin=subprocess.DEVNULL,
                    stdout=log, stderr=log, start_new_session=True))
        urls = (config["service_url"] + "/health", config["ui_url"] + "/_stcore/health")
        for url in urls:
            for _ in range(50):
                if any(process.poll() is not None for process in processes):
                    raise RuntimeError("A service exited. See .figma_bridge logs.")
                try:
                    with urlopen(url, timeout=1) as response:
                        if response.status == 200:
                            break
                except OSError:
                    pass
                time.sleep(0.1)
            else:
                raise RuntimeError("Local service did not become ready.")
        config["pids"] = [process.pid for process in processes]
        with os.fdopen(os.open(config_path, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600), "w") as output:
            json.dump(config, output, indent=2)
        print(f"Local handoff: {config['ui_url']}")
        print(f"Plugin service URL: {config['service_url']}")
    except Exception:
        for process in processes:
            process.terminate()
            process.wait(timeout=5)
        raise


if __name__ == "__main__":
    main()
