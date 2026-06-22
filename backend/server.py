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

_APP_USER    = os.getenv("APP_USER", "team").strip()
_APP_PASS    = os.getenv("APP_PASSWORD", "").strip()
_SECRET_KEY  = os.getenv("SECRET_KEY", secrets.token_hex(32)).strip()
_AUTOMACAO_API_KEY = os.getenv("AUTOMACAO_ML_API_KEY", "").strip()

print(f"\n--- CONFIGURAÇÃO DE ACESSO ---")
print(f"Usuário: '{_APP_USER}'")
if not _APP_PASS:
    print("Senha: NÃO CONFIGURADA (Acesso livre)")
else:
    print("Senha: CONFIGURADA")
print("API Key Auth: CONFIGURADA" if _AUTOMACAO_API_KEY else "API Key Auth: DESATIVADA")
print(f"------------------------------\n")

_jobs: dict[str, queue.Queue] = {}
_running = False
_proc: "subprocess.Popen | None" = None
_cancel_requested = False


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


def require_session(request: Request, session: str | None = Cookie(default=None)):
    # Permite acesso via API Key (chamadas do sistemaMelo)
    if _AUTOMACAO_API_KEY:
        header_key = request.headers.get("X-API-Key", "")
        if secrets.compare_digest(header_key.encode(), _AUTOMACAO_API_KEY.encode()):
            return
    # Fallback: auth por cookie (acesso direto ao frontend próprio)
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
  <title>Melo Midia — Automação de Textos</title>
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet"/>
  <style>
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: 'Inter', system-ui, sans-serif;
      background-color: #07070f;
      color: #e2e8f0;
      min-height: 100vh;
      display: flex;
      align-items: center;
      justify-content: center;
      overflow: hidden;
      background-image:
        radial-gradient(ellipse 80% 50% at 20% -10%, rgba(124, 58, 237, 0.15) 0%, transparent 60%),
        radial-gradient(ellipse 60% 40% at 80% 110%, rgba(6, 182, 212, 0.1) 0%, transparent 60%);
    }
    /* Animated Orbs */
    .orb {
      position: absolute;
      border-radius: 50%;
      filter: blur(80px);
      z-index: -1;
      animation: float 20s infinite ease-in-out alternate;
    }
    .orb-1 {
      width: 400px; height: 400px;
      background: rgba(124, 58, 237, 0.15);
      top: -100px; left: -100px;
    }
    .orb-2 {
      width: 300px; height: 300px;
      background: rgba(6, 182, 212, 0.1);
      bottom: -50px; right: -50px;
      animation-delay: -5s;
    }
    @keyframes float {
      0% { transform: translate(0, 0) scale(1); }
      100% { transform: translate(50px, 30px) scale(1.1); }
    }
    
    .card {
      background: rgba(16, 16, 30, 0.6);
      border: 1px solid rgba(255, 255, 255, 0.08);
      border-radius: 20px;
      padding: 48px 40px;
      width: 100%;
      max-width: 420px;
      backdrop-filter: blur(24px);
      -webkit-backdrop-filter: blur(24px);
      box-shadow: 0 24px 64px rgba(0, 0, 0, 0.5), inset 0 1px 0 rgba(255, 255, 255, 0.05);
      animation: card-in 0.6s cubic-bezier(0.22, 1, 0.36, 1);
      position: relative;
    }
    @keyframes card-in {
      from { opacity: 0; transform: translateY(20px) scale(0.98); }
      to   { opacity: 1; transform: translateY(0) scale(1); }
    }
    
    .logo { text-align: center; margin-bottom: 40px; }
    .logo-icon {
      width: 72px; height: 72px;
      border-radius: 12px;
      display: inline-flex; align-items: center; justify-content: center;
      margin-bottom: 16px;
      box-shadow: 0 0 24px rgba(124, 58, 237, 0.2);
      position: relative;
      overflow: hidden;
    }
    .logo-icon img {
      width: 100%;
      height: 100%;
      object-fit: cover;
    }
    .logo-icon::after {
      content: '';
      position: absolute; inset: 0;
      border-radius: inherit;
      border: 1px solid rgba(255, 255, 255, 0.1);
    }
    .logo-title { display: block; font-size: 1.5rem; font-weight: 700; letter-spacing: -0.02em; }
    .logo-sub { display: block; font-size: 0.85rem; color: #94a3b8; margin-top: 6px; font-weight: 400; }
    
    .input-group { position: relative; margin-bottom: 20px; }
    label { 
      position: absolute;
      left: 16px; top: 50%;
      transform: translateY(-50%);
      font-size: 0.9rem; font-weight: 400; color: #64748b;
      pointer-events: none;
      transition: all 0.2s ease;
      background: transparent;
      padding: 0 4px;
    }
    input[type=text], input[type=password] {
      width: 100%;
      background: rgba(0, 0, 0, 0.2);
      border: 1px solid rgba(255, 255, 255, 0.1);
      border-radius: 12px;
      color: #e2e8f0;
      font-size: 0.95rem;
      padding: 16px;
      outline: none;
      transition: all 0.25s ease;
      font-family: inherit;
    }
    input:focus, input:not(:placeholder-shown) {
      background: rgba(0, 0, 0, 0.4);
    }
    input:focus {
      border-color: #7c3aed;
      box-shadow: 0 0 0 4px rgba(124, 58, 237, 0.15);
    }
    input:focus + label, input:not(:placeholder-shown) + label {
      top: 0; transform: translateY(-50%) scale(0.85);
      background: #11111d; /* Matches input border intersection */
      color: #9d5cf6;
      border-radius: 4px;
    }
    input:not(:focus):not(:placeholder-shown) + label {
      color: #94a3b8;
    }
    
    button {
      width: 100%;
      background: linear-gradient(135deg, #7c3aed, #6d28d9);
      border: 1px solid rgba(255,255,255,0.1);
      border-radius: 12px;
      color: #fff;
      font-size: 1rem;
      font-weight: 600;
      padding: 16px;
      cursor: pointer;
      margin-top: 12px;
      transition: all 0.3s cubic-bezier(0.22, 1, 0.36, 1);
      box-shadow: 0 4px 20px rgba(124, 58, 237, 0.3);
      position: relative;
      overflow: hidden;
    }
    button::after {
      content: '';
      position: absolute; top: 0; left: -100%; width: 50%; height: 100%;
      background: linear-gradient(90deg, transparent, rgba(255,255,255,0.15), transparent);
      transform: skewX(-20deg);
      transition: 0.5s;
    }
    button:hover {
      transform: translateY(-2px);
      box-shadow: 0 8px 28px rgba(124, 58, 237, 0.4);
    }
    button:hover::after { left: 150%; }
    button:active { transform: translateY(0); box-shadow: 0 4px 16px rgba(124, 58, 237, 0.3); }
    
    .error {
      background: rgba(239, 68, 68, 0.1);
      border: 1px solid rgba(239, 68, 68, 0.3);
      border-radius: 12px;
      color: #fca5a5;
      font-size: 0.85rem;
      padding: 12px 16px;
      margin-bottom: 24px;
      display: flex; align-items: center; gap: 8px;
      animation: shake 0.5s cubic-bezier(0.36, 0.07, 0.19, 0.97) both;
    }
    @keyframes shake {
      10%, 90% { transform: translate3d(-1px, 0, 0); }
      20%, 80% { transform: translate3d(2px, 0, 0); }
      30%, 50%, 70% { transform: translate3d(-4px, 0, 0); }
      40%, 60% { transform: translate3d(4px, 0, 0); }
    }
  </style>
</head>
<body>
  <div class="orb orb-1"></div>
  <div class="orb orb-2"></div>

  <div class="card">
    <div class="logo">
      <div class="logo-icon">
        <img src="/static/logo.png" alt="Melo Midia Logo"/>
      </div>
      <span class="logo-title">Melo Midia</span>
      <span class="logo-sub">Automação de Textos</span>
    </div>
    
    {error_block}
    
    <form method="post" action="/login">
      <div class="input-group">
        <input type="text" id="u" name="username" autocomplete="username" placeholder=" " required/>
        <label for="u">Usuário</label>
      </div>
      <div class="input-group">
        <input type="password" id="p" name="password" autocomplete="current-password" placeholder=" " required/>
        <label for="p">Senha</label>
      </div>
      <button type="submit">Entrar no Sistema</button>
    </form>
  </div>
</body>
</html>"""


@app.get("/login", response_class=HTMLResponse)
def login_page():
    return LOGIN_HTML.replace("{error_block}", "")


@app.post("/login")
def login(username: str = Form(...), password: str = Form(...)):
    # Username check (case-insensitive for convenience)
    ok_user = secrets.compare_digest(username.strip().lower().encode(), _APP_USER.lower().encode())
    # Password check (case-sensitive)
    ok_pass = secrets.compare_digest(password.encode(), _APP_PASS.encode()) if _APP_PASS else True
    
    if not (ok_user and ok_pass):
        print(f"ALERTA: Tentativa de login falhou.")
        print(f"  - Usuário fornecido: '{username}'")
        print(f"  - Usuário esperado: '{_APP_USER}'")
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


@app.get("/api/sheets")
def get_sheets(client_id: str, _: None = Depends(require_session)):
    try:
        creds  = aml.get_credentials()
        drive  = build("drive",  "v3", credentials=creds)
        sheets = build("sheets", "v4", credentials=creds)
        spreadsheet = aml.find_main_spreadsheet(drive, client_id)
        if not spreadsheet:
            return {"ok": True, "sheets": []}
        names = aml.get_sheet_names(drive, sheets, spreadsheet)
        return {"ok": True, "sheets": names}
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
        global _running, _proc, _cancel_requested
        _running = True
        _cancel_requested = False
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
            _proc = proc
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

            try:
                proc.wait(timeout=5)
            except subprocess.TimeoutExpired:
                proc.kill()
                proc.wait()
            except Exception:
                pass

            if not done_received:
                if _cancel_requested:
                    q.put({"type": "cancelled"})
                else:
                    q.put({"type": "error", "message": "Processo encerrou inesperadamente."})

        except Exception as exc:
            for line in traceback.format_exc().splitlines():
                q.put({"type": "log", "text": line})
            q.put({"type": "error", "message": str(exc)})
        finally:
            _proc = None
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
                if event.get("type") in ("done", "error", "cancelled"):
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


@app.post("/api/cancel")
def cancel_automation(_: None = Depends(require_session)):
    global _cancel_requested, _proc
    _cancel_requested = True
    if _proc is not None and _proc.poll() is None:
        _proc.kill()  # SIGKILL no Linux — não pode ser ignorado
    return {"ok": True}


@app.post("/api/reset")
def reset_running(_: None = Depends(require_session)):
    global _running
    _running = False
    return {"ok": True}


if __name__ == "__main__":
    uvicorn.run("server:app", host="127.0.0.1", port=8000, reload=False)
