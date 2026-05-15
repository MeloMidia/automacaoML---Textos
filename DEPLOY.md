# Deploy no Railway

## Pré-requisitos
- Conta no [Railway](https://railway.app) (gratuito)
- Repositório no GitHub com o código

---

## Passo 1 — Gerar o TOKEN_PICKLE_B64

Execute esse comando **uma vez** na sua máquina local (onde o `token.pickle` já existe):

**Windows (PowerShell):**
```powershell
[Convert]::ToBase64String([IO.File]::ReadAllBytes("token.pickle"))
```

Copie o resultado — você vai colar no Railway como variável de ambiente.

---

## Passo 2 — Subir para o GitHub

1. Crie um repositório no GitHub
2. Faça push de todos os arquivos (o `token.pickle` e `credentials.json` **não** precisam ir — as variáveis de ambiente substituem)

Crie um `.gitignore` com:
```
token.pickle
credentials.json
__pycache__/
*.pyc
.env
```

---

## Passo 3 — Deploy no Railway

1. Acesse [railway.app](https://railway.app) → **New Project** → **Deploy from GitHub repo**
2. Selecione o repositório
3. Vá em **Variables** e adicione:

| Variável | Valor |
|---|---|
| `GROQ_API_KEY` | Sua chave do Groq |
| `APP_USER` | `team` (ou o nome que quiser) |
| `APP_PASSWORD` | A senha que o time vai usar |
| `TOKEN_PICKLE_B64` | O valor gerado no Passo 1 |

4. O Railway detecta o `Dockerfile` automaticamente e faz o deploy
5. Copie a URL gerada (ex: `https://automacaoml.up.railway.app`) e compartilhe com o time

---

## Acessar a aplicação

- Abra a URL no navegador
- Digite o usuário e senha configurados
- Pronto — funciona igual ao local, de qualquer computador
