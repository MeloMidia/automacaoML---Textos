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

from dotenv import load_dotenv
load_dotenv(Path(__file__).resolve().parent.parent / ".env")

import uvicorn
from fastapi import Cookie, Depends, FastAPI, Form, Request, status
from fastapi.responses import FileResponse, HTMLResponse, RedirectResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from googleapiclient.discovery import build

ROOT         = Path(__file__).resolve().parent.parent
RUNNER       = Path(__file__).resolve().parent / "runner.py"
FRONTEND_DIR = ROOT / "frontend"

sys.path.insert(0, str(ROOT))
import automacao_ml as aml  # noqa

app = FastAPI(title="AutomacaoML Web")
app.mount("/static", StaticFiles(directory=str(FRONTEND_DIR)), name="static")

_APP_USER    = os.getenv("APP_USER", "team")
_APP_PASS    = os.getenv("APP_PASSWORD", "")
_SECRET_KEY  = os.getenv("SECRET_KEY", secrets.token_hex(32))

_jobs: dict[str, queue.Queue] = {}
_running = False


# ── Sessão simples via cookie assinado ─────────────────────────────────────

def _make_token(username: str) -> str:
    import hmac, hashlib
    msg = f"{username}:{_SECRET_KEY}".encode()
    sig = hmac.new(_SECRET_KEY.encode(), msg, hashlib.sha256).hexdigest()
    return f"{username}:{sig}"


def _verify_token(token: str) -> bool:
    try:
        username, sig = token.rsplit(":", 1)
        import hmac, hashlib
        msg = f"{username}:{_SECRET_KEY}".encode()
        expected = hmac.new(_SECRET_KEY.encode(), msg, hashlib.sha256).hexdigest()
        return secrets.compare_digest(sig, expected)
    except Exception:
        return False


def require_session(session: str | None = Cookie(default=None)):
    if not _APP_PASS:
        return  # sem senha configurada, permite acesso (útil em dev local)
    if not session or not _verify_token(session):
        raise RedirectException("/login")


class RedirectException(Exception):
    def __init__(self, url: str):
        self.url = url


@app.exception_handler(RedirectException)
async def redirect_handler(request: Request, exc: RedirectException):
    return RedirectResponse(url=exc.url, status_code=status.HTTP_302_FOUND)


# ── Login ───────────────────────────────────────────────────────────────────

LOGIN_HTML = """<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <title>AutomacaoML — Login</title>
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&display=swap" rel="stylesheet"/>
  <style>
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: 'Inter', sans-serif;
      background: #0f1117;
      color: #e2e8f0;
      min-height: 100vh;
      display: flex;
      align-items: center;
      justify-content: center;
    }
    .card {
      background: #1a1d27;
      border: 1px solid #2d3148;
      border-radius: 16px;
      padding: 40px;
      width: 100%;
      max-width: 380px;
    }
    .logo { text-align: center; margin-bottom: 32px; }
    .logo-icon {
      width: 52px; height: 52px;
      background: linear-gradient(135deg, #6366f1, #8b5cf6);
      border-radius: 14px;
      display: inline-flex; align-items: center; justify-content: center;
      margin-bottom: 14px;
    }
    .logo-title { display: block; font-size: 1.25rem; font-weight: 600; }
    .logo-sub { display: block; font-size: 0.8rem; color: #64748b; margin-top: 4px; }
    label { display: block; font-size: 0.8rem; font-weight: 500; color: #94a3b8; margin-bottom: 6px; }
    input[type=text], input[type=password] {
      width: 100%;
      background: #0f1117;
      border: 1px solid #2d3148;
      border-radius: 8px;
      color: #e2e8f0;
      font-size: 0.9rem;
      padding: 10px 14px;
      outline: none;
      margin-bottom: 16px;
      transition: border-color .2s;
    }
    input:focus { border-color: #6366f1; }
    button {
      width: 100%;
      background: linear-gradient(135deg, #6366f1, #8b5cf6);
      border: none;
      border-radius: 8px;
      color: #fff;
      font-size: 0.95rem;
      font-weight: 600;
      padding: 12px;
      cursor: pointer;
      margin-top: 8px;
      transition: opacity .2s;
    }
    button:hover { opacity: .9; }
    .error {
      background: #2d1b1b;
      border: 1px solid #7f1d1d;
      border-radius: 8px;
      color: #fca5a5;
      font-size: 0.82rem;
      padding: 10px 14px;
      margin-bottom: 16px;
    }
  </style>
</head>
<body>
  <div class="card">
    <div class="logo">
      <div class="logo-icon">
        <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="white" stroke-width="2">
          <path d="M13 2L3 14h9l-1 8 10-12h-9l1-8z"/>
        </svg>
      </div>
      <span class="logo-title">AutomacaoML</span>
      <span class="logo-sub">Gerador de Anúncios via IA</span>
    </div>
    {error_block}
    <form method="post" action="/login">
      <label for="u">Usuário</label>
      <input type="text" id="u" name="username" autocomplete="username" required/>
      <label for="p">Senha</label>
      <input type="password" id="p" name="password" autocomplete="current-password" required/>
      <button type="submit">Entrar</button>
    </form>
  </div>
</body>
</html>"""


@app.get("/login", response_class=HTMLResponse)
def login_page():
    return LOGIN_HTML.replace("{error_block}", "")


@app.post("/login")
def login(username: str = Form(...), password: str = Form(...)):
    ok_user = secrets.compare_digest(username.encode(), _APP_USER.encode())
    ok_pass = secrets.compare_digest(password.encode(), _APP_PASS.encode()) if _APP_PASS else True
    if not (ok_user and ok_pass):
        error = '<div class="error">Usuário ou senha incorretos.</div>'
        return HTMLResponse(LOGIN_HTML.replace("{error_block}", error), status_code=401)
    response = RedirectResponse(url="/", status_code=status.HTTP_302_FOUND)
    response.set_cookie("session", _make_token(username), httponly=True, samesite="lax")
    return response


@app.get("/logout")
def logout():
    response = RedirectResponse(url="/login", status_code=status.HTTP_302_FOUND)
    response.delete_cookie("session")
    return response


# ── Endpoints protegidos ────────────────────────────────────────────────────

@app.get("/")
def serve_index(_: None = Depends(require_session)):
    return FileResponse(str(FRONTEND_DIR / "index.html"))


@app.get("/api/clients")
def get_clients(_: None = Depends(require_session)):
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
def run_automation(body: dict, _: None = Depends(require_session)):
    global _running
    if _running:
        return {"ok": False, "error": "Já existe uma automação em execução."}

    selected = body.get("clients", [])
    if not selected:
        return {"ok": False, "error": "Nenhum cliente selecionado."}

    delay  = max(10, int(body.get("delay_seconds", 45)))
    job_id = str(uuid.uuid4())
    q: queue.Queue = queue.Queue()
    _jobs[job_id] = q

    def worker():
        global _running
        _running = True
        try:
            config_json = json.dumps({"clients": selected, "delay": delay})
            env = {**__import__("os").environ, "PYTHONUNBUFFERED": "1", "PYTHONUTF8": "1"}
            proc = subprocess.Popen(
                [sys.executable, str(RUNNER)],
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
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
def stream_logs(job_id: str, _: None = Depends(require_session)):
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
def get_status(_: None = Depends(require_session)):
    return {"running": _running}


if __name__ == "__main__":
    uvicorn.run("server:app", host="127.0.0.1", port=8000, reload=False)
