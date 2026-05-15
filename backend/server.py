#!/usr/bin/env python3
"""
AutomacaoML — Servidor Web (FastAPI)
Inicie com: python -m uvicorn backend.server:app --host 127.0.0.1 --port 8000
"""

import json
import os
import queue
import secrets
import subprocess
import sys
import threading
import traceback
import uuid
from pathlib import Path

import uvicorn
from fastapi import Depends, FastAPI, HTTPException, status
from fastapi.security import HTTPBasic, HTTPBasicCredentials
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from googleapiclient.discovery import build

ROOT         = Path(__file__).resolve().parent.parent
RUNNER       = Path(__file__).resolve().parent / "runner.py"
FRONTEND_DIR = ROOT / "frontend"

sys.path.insert(0, str(ROOT))
import automacao_ml as aml  # noqa

app = FastAPI(title="AutomacaoML Web")
app.mount("/static", StaticFiles(directory=str(FRONTEND_DIR)), name="static")

_security = HTTPBasic()
_APP_USER = os.getenv("APP_USER", "team")
_APP_PASS = os.getenv("APP_PASSWORD", "")

_jobs: dict[str, queue.Queue] = {}
_running = False


def require_auth(credentials: HTTPBasicCredentials = Depends(_security)):
    if not _APP_PASS:
        return  # sem senha configurada, permite acesso (útil em dev local)
    ok_user = secrets.compare_digest(credentials.username.encode(), _APP_USER.encode())
    ok_pass = secrets.compare_digest(credentials.password.encode(), _APP_PASS.encode())
    if not (ok_user and ok_pass):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Usuário ou senha incorretos.",
            headers={"WWW-Authenticate": "Basic"},
        )


# ── Endpoints ──────────────────────────────────────────────────────────────

@app.get("/")
def serve_index(_: None = Depends(require_auth)):
    return FileResponse(str(FRONTEND_DIR / "index.html"))


@app.get("/api/clients")
def get_clients(_: None = Depends(require_auth)):
    try:
        creds   = aml.get_credentials()
        drive   = build("drive", "v3", credentials=creds)
        clients = aml.list_subfolders(drive, aml.PASTA_RAIZ_ID)
        return {"ok": True, "clients": clients}
    except SystemExit:
        return {"ok": False, "error": "credentials.json não encontrado."}
    except Exception as e:
        return {"ok": False, "error": str(e)}


@app.post("/api/run")
def run_automation(body: dict, _: None = Depends(require_auth)):
    global _running
    if _running:
        return {"ok": False, "error": "Já existe uma automação em execução."}

    selected = body.get("clients", [])
    if not selected:
        return {"ok": False, "error": "Nenhum cliente selecionado."}

    delay    = max(10, int(body.get("delay_seconds", 45)))
    job_id   = str(uuid.uuid4())
    q: queue.Queue = queue.Queue()
    _jobs[job_id] = q

    def worker():
        global _running
        _running = True
        try:
            config_json = json.dumps({"clients": selected, "delay": delay})

            # ── Processo isolado: zero problemas de thread/stdout no Windows ──
            env = {**__import__("os").environ, "PYTHONUNBUFFERED": "1", "PYTHONUTF8": "1"}
            proc = subprocess.Popen(
                [sys.executable, str(RUNNER)],
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,   # stderr redireciona para stdout
                text=True,
                encoding="utf-8",
                errors="replace",
                cwd=str(ROOT),
                env=env,
            )

            proc.stdin.write(config_json)
            proc.stdin.close()

            done_received = False
            for raw in proc.stdout:
                line = raw.rstrip("\n")
                if not line:
                    continue
                try:
                    event = json.loads(line)
                    q.put(event)
                    if event.get("type") in ("done", "error"):
                        done_received = True
                        break
                except json.JSONDecodeError:
                    # Linha de print() normal do script → trata como log
                    q.put({"type": "log", "text": line})

            proc.wait(timeout=5)

            if not done_received:
                q.put({"type": "error", "message": "Processo encerrou inesperadamente."})

        except Exception as exc:
            for line in traceback.format_exc().splitlines():
                q.put({"type": "log", "text": line})
            q.put({"type": "error", "message": str(exc)})
        finally:
            _running = False

    threading.Thread(target=worker, daemon=True).start()
    return {"ok": True, "job_id": job_id}


@app.get("/api/stream/{job_id}")
def stream_logs(job_id: str, _: None = Depends(require_auth)):
    q = _jobs.get(job_id)
    if not q:
        def _nf():
            yield f"data: {json.dumps({'type': 'error', 'message': 'Job não encontrado.'})}\n\n"
        return StreamingResponse(_nf(), media_type="text/event-stream")

    def generate():
        while True:
            try:
                event = q.get(timeout=25)
                yield f"data: {json.dumps(event, ensure_ascii=False)}\n\n"
                if event.get("type") in ("done", "error"):
                    _jobs.pop(job_id, None)
                    break
            except queue.Empty:
                yield f"data: {json.dumps({'type': 'ping'})}\n\n"

    return StreamingResponse(
        generate(),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


@app.get("/api/status")
def get_status(_: None = Depends(require_auth)):
    return {"running": _running}


if __name__ == "__main__":
    uvicorn.run("server:app", host="127.0.0.1", port=8000, reload=False)
