"""Converte token.pickle para base64 e salva em token_b64.txt."""
import base64
from pathlib import Path

token_path = Path("token.pickle")
if not token_path.exists():
    print("ERRO: token.pickle não encontrado. Rode o script principal primeiro.")
    raise SystemExit(1)

b64 = base64.b64encode(token_path.read_bytes()).decode()
Path("token_b64.txt").write_text(b64, encoding="utf-8")
print(f"token_b64.txt gerado ({len(b64)} chars).")
print("Copie o conteúdo de token_b64.txt para a variável TOKEN_PICKLE_B64 no Render.")
