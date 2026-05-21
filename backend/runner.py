#!/usr/bin/env python3
"""
Processo isolado para execução da automação.
Recebe JSON via stdin, emite JSON por linha no stdout.
"""
import io, json, sys, traceback
from pathlib import Path

# Forca UTF-8 no stdout do subprocess (Windows usa cp1252 por padrao)
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
else:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import automacao_ml as aml
from googleapiclient.discovery import build


def emit(type_, **kwargs):
    # ensure_ascii=True: saida 100% ASCII, safe em qualquer encoding de terminal
    print(json.dumps({"type": type_, **kwargs}, ensure_ascii=True), flush=True)


def main():
    config   = json.loads(sys.stdin.read())
    selected = config.get("clients", [])
    delay    = max(10, int(config.get("delay", 45)))

    aml.DELAY_ENTRE_PRODUTOS = delay
    emit("log", text=f"⚙️  Delay: {delay}s por produto")

    emit("log", text="🔑 Autenticando com Google...")
    creds = aml.get_credentials()
    emit("log", text="✅ Autenticado!")

    drive  = build("drive",  "v3", credentials=creds)
    sheets = build("sheets", "v4", credentials=creds)
    docs   = build("docs",   "v1", credentials=creds)

    total_c = total_s = 0
    for client in selected:
        sheets_filter = client.get("sheets") or None
        c, s = aml.process_client(client["name"], client["id"], drive, sheets, docs, sheets_filter=sheets_filter)
        total_c += c
        total_s += s

    emit("done", created=total_c, skipped=total_s)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        for line in traceback.format_exc().splitlines():
            emit("log", text=line)
        emit("error", message=str(exc))
        sys.exit(1)
