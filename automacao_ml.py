#!/usr/bin/env python3
"""
============================================================
  AUTOMAÇÃO MERCADO LIVRE — Geração de Anúncios via Gemini
============================================================
Lê produtos da planilha no Google Drive, gera título e
descrição via API do Gemini e salva como Google Docs na
pasta TEXTOS do cliente.

SETUP RÁPIDO:
  1. pip install -r requirements.txt
  2. Baixe credentials.json do Google Cloud Console
     (veja INSTRUCOES.txt para o passo a passo)
  3. python automacao_ml.py
============================================================
"""

import base64
import io
import os
import pickle
import re
import time
from pathlib import Path

# Força o diretório de trabalho para a pasta onde este script está salvo
os.chdir(Path(__file__).resolve().parent)

import openpyxl
from groq import Groq
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# ─────────────────────────────────────────────────────────
#  CONFIGURAÇÕES — edite aqui se necessário
# ─────────────────────────────────────────────────────────

GROQ_API_KEY = os.getenv("GROQ_API_KEY", "")

# ID da pasta MERCADO LIVRE no Google Drive (fixo — não precisa alterar)
PASTA_RAIZ_ID = "1H7r7kvGIuuqZByHaAVXxvkHA64852_if"

# Linha de início dos dados na planilha (pula cabeçalhos)
LINHA_INICIO = 3  # linha 3 = primeira linha de produto

# Índices das colunas (0 = coluna A, 1 = B, etc.)
# Ordem da planilha: SKU(A), PRODUTO(B), MARCA(C), CÓDIGO(D), APLICAÇÃO(E)
COL_PRODUTO   = 1   # B — Nome do produto
COL_MARCA     = 2   # C — Marca
COL_CODIGO    = 3   # D — Código
COL_APLICACAO = 4   # E — Aplicação

# Intervalo entre chamadas à API (segundos) para evitar rate limit
DELAY_ENTRE_PRODUTOS = 20

# ─────────────────────────────────────────────────────────

SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/documents",
]


# ══════════════════════════════════════════════════════════
#  AUTENTICAÇÃO
# ══════════════════════════════════════════════════════════

def get_credentials():
    """Autentica com Google OAuth2.

    Em servidor (deploy): lê o token de TOKEN_PICKLE_B64 (base64 do token.pickle).
    Localmente: comportamento original com arquivo token.pickle.
    """
    creds = None
    credentials_path = Path("credentials.json")
    token_path = Path("token.pickle")

    token_b64 = os.getenv("TOKEN_PICKLE_B64", "")
    if token_b64:
        creds = pickle.loads(base64.b64decode(token_b64))
    elif token_path.exists():
        with open(token_path, "rb") as f:
            creds = pickle.load(f)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
            if not token_b64:
                with open(token_path, "wb") as f:
                    pickle.dump(creds, f)
        else:
            if not credentials_path.exists():
                raise SystemExit(
                    "credentials.json não encontrado e TOKEN_PICKLE_B64 não definido."
                )
            flow = InstalledAppFlow.from_client_secrets_file(str(credentials_path), SCOPES)
            creds = flow.run_local_server(port=0)
            with open(token_path, "wb") as f:
                pickle.dump(creds, f)

    return creds


# ══════════════════════════════════════════════════════════
#  FUNÇÕES DO GOOGLE DRIVE
# ══════════════════════════════════════════════════════════

def find_folder_by_name(drive, name, parent_id):
    """Busca uma pasta pelo nome dentro de uma pasta pai."""
    q = f"mimeType='application/vnd.google-apps.folder' and name='{name}' and '{parent_id}' in parents and trashed=false"
    r = drive.files().list(q=q, fields="files(id,name)").execute()
    files = r.get("files", [])
    return files[0] if files else None


def list_subfolders(drive, parent_id):
    """Lista todas as subpastas de uma pasta."""
    q = f"mimeType='application/vnd.google-apps.folder' and '{parent_id}' in parents and trashed=false"
    r = drive.files().list(q=q, fields="files(id,name)", orderBy="name").execute()
    return r.get("files", [])


def find_main_spreadsheet(drive, parent_id):
    """Encontra a planilha principal na pasta do cliente."""
    q = (
        f"(mimeType='application/vnd.google-apps.spreadsheet' or "
        f"mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') "
        f"and '{parent_id}' in parents and trashed=false"
    )
    r = drive.files().list(q=q, fields="files(id,name,mimeType)").execute()
    files = r.get("files", [])
    if not files:
        return None
    # Prefere arquivos que não sejam "PLANILHA DANI" (ajuste se precisar)
    return files[0]


def get_or_create_textos_folder(drive, client_folder_id):
    """Retorna (ou cria) a pasta TEXTOS dentro do cliente."""
    folder = find_folder_by_name(drive, "TEXTOS", client_folder_id)
    if folder:
        return folder["id"]
    meta = {
        "name": "TEXTOS",
        "mimeType": "application/vnd.google-apps.folder",
        "parents": [client_folder_id],
    }
    new_folder = drive.files().create(body=meta, fields="id").execute()
    print("     📁 Pasta TEXTOS criada.")
    return new_folder["id"]


def list_existing_docs(drive, folder_id):
    """Retorna um set com os nomes dos arquivos já existentes em TEXTOS."""
    q = f"'{folder_id}' in parents and trashed=false"
    r = drive.files().list(q=q, fields="files(name)").execute()
    return {f["name"] for f in r.get("files", [])}


# ══════════════════════════════════════════════════════════
#  LEITURA DA PLANILHA
# ══════════════════════════════════════════════════════════

def read_products_from_spreadsheet(drive, sheets, file_info):
    """Lê os produtos da planilha (Google Sheets nativo ou .xlsx)."""
    fid   = file_info["id"]
    mime  = file_info["mimeType"]

    if mime == "application/vnd.google-apps.spreadsheet":
        # ── Google Sheets nativo ──────────────────────────
        result = sheets.spreadsheets().values().get(
            spreadsheetId=fid,
            range=f"A{LINHA_INICIO}:Z2000"
        ).execute()
        rows = result.get("values", [])
    else:
        # ── Arquivo .xlsx ─────────────────────────────────
        request = drive.files().get_media(fileId=fid)
        buf = io.BytesIO()
        dl = MediaIoBaseDownload(buf, request)
        done = False
        while not done:
            _, done = dl.next_chunk()
        buf.seek(0)
        wb = openpyxl.load_workbook(buf, data_only=True)
        ws = wb.active
        rows = []
        for row in ws.iter_rows(min_row=LINHA_INICIO, values_only=True):
            rows.append([str(c) if c is not None else "" for c in row])

    products = []
    for row in rows:
        if len(row) <= COL_PRODUTO:
            continue

        def cell(idx):
            if idx is None or len(row) <= idx:
                return ""
            v = str(row[idx]).strip()
            return "" if v.lower() in ("none", "nan", "") else v

        produto   = cell(COL_PRODUTO)
        marca     = cell(COL_MARCA)
        aplicacao = cell(COL_APLICACAO)
        codigo    = cell(COL_CODIGO) or "SEM"

        if produto and produto.lower() not in ("produto", "none", "nan"):
            products.append({
                "produto":   produto,
                "marca":     marca,
                "aplicacao": aplicacao,
                "codigo":    codigo or "SEM",
            })

    return products


# ══════════════════════════════════════════════════════════
#  GERAÇÃO VIA GEMINI
# ══════════════════════════════════════════════════════════

def generate_with_groq(product, client_name):
    """Envia o prompt ao Groq e retorna o texto gerado, com retry automático."""
    client = Groq(api_key=GROQ_API_KEY)

    system_message = """Você é um especialista em SEO e copywriting para o Mercado Livre, criando anúncios de alta conversão para produtos industriais, automotivos e de abastecimento.

REGRAS ABSOLUTAS — NUNCA QUEBRE:
1. Responda SOMENTE com o conteúdo do anúncio. Zero introduções, zero "Aqui está", zero comentários.
2. NUNCA use asteriscos (*), hashtags (#), aspas ou qualquer markdown.
3. O título é uma frase SEO. NUNCA copie o nome bruto do produto como título.
4. Escreva o título APENAS UMA VEZ, na linha após o rótulo "TITULO".
5. NÃO existe seção VARIACAO. O anúncio tem apenas UM título.
6. Nos dois parágrafos de texto, NUNCA use "nosso", "nossa", "nós". Sempre terceira pessoa.
7. Na seção APLICACAO, use SOMENTE os dados fornecidos. NUNCA invente compatibilidades.
8. Os dois parágrafos de texto devem ser técnicos e específicos ao produto. NUNCA genéricos."""

    aplicacao = product['aplicacao'] if product['aplicacao'] else ''

    prompt = f"""Crie um anúncio completo para o Mercado Livre. Siga o modelo abaixo EXATAMENTE.

========================================================
EXEMPLO REAL DE ANUNCIO BEM FEITO (use como referência de qualidade):
========================================================

DADOS DE ENTRADA DO EXEMPLO:
Produto: Mangote Para Breakaway 3/4 150MM Fixo + Fixo Alumínio
Marca: MARTINELLI
Codigo: SEM
Aplicacao: Utilizado para instalação da válvula de segurança Breakaway em bombas de combustível / Compatível com sistemas de abastecimento 3/4 polegadas / Entrada e saída macho 3/4" NPT / Indicado para postos de combustíveis

SAIDA CORRETA DO EXEMPLO:

TITULO
Mangote Breakaway 3/4 150mm Fixo Fixo Alumínio

DESCRICAO COMPLETA 
--------------------------------------------------
O QUE VEM NA CAIXA
- 01 Mangote Para Breakaway 3/4 150MM Fixo + Fixo Alumínio
- Marca: MARTINELLI
- Codigo/Referencia: SEM
--------------------------------------------------
APLICACAO
- Utilizado para instalação da válvula de segurança Breakaway
- Aplicação em bombas de combustível
- Compatível com sistemas de abastecimento 3/4 polegadas
- Entrada e saída macho 3/4" NPT
- Indicado para postos de combustíveis e sistemas de abastecimento
--------------------------------------------------
O mangote para breakaway 3/4 150mm fixo + fixo alumínio é um componente essencial para a instalação correta da válvula de segurança Breakaway em bombas de combustível, garantindo uma conexão segura e eficiente no sistema de abastecimento.

Fabricado com mangueira 3/4" com trama de aço, oferece maior resistência e durabilidade para operações contínuas, utilizando mangueira certificada e homologada pelo INMETRO. Conta com dois terminais fixos de alumínio de alta qualidade e conexões macho 3/4" NPT, proporcionando vedação eficiente, resistência ao uso intenso e maior confiabilidade na operação do posto.
--------------------------------------------------
INSTITUCIONAL

A MUNDIAL POSTO oferece produtos de qualidade para abastecimento, manutenção e operação de postos de combustíveis, disponibilizando soluções confiáveis para maior segurança e eficiência no dia a dia.

Nosso compromisso é entregar produtos de procedência, envio rápido e atendimento de confiança, ajudando sua operação a manter desempenho, organização e segurança.

========================================================
REGRAS DO TITULO (aprenda com o exemplo):
- O título não é o nome do produto. É uma frase de busca otimizada.
- Inclua especificações técnicas e/ou modelos compatíveis extraídos da APLICACAO.
- Máximo 60 caracteres. Sem a marca. Sem o código.
- RUIM: "LAMPADA H7 24V 100W CAMINHAO"  (nome bruto)
- BOM: "Lampada H7 24V 100W Caminhao Truck Par Alta Potencia"  (especificações + contexto de busca)
- RUIM: "FILTRO COMBUSTIVEL"  (genérico demais)
- BOM: "Filtro Combustivel Gol Voyage Polo Clio Fit Civic 1.0 1.4"  (modelos compatíveis)

========================================================
REGRAS DA SECAO APLICACAO (aprenda com o exemplo):
- Para produtos de posto/industrial: liste usos, compatibilidades e especificações técnicas como bullets.
- Para autopeças com lista de veículos: agrupe por montadora:
  Volkswagen: Gol, Voyage, Polo, Golf
  Fiat: Uno, Palio, Siena
  Honda: Civic, Fit
- Se a aplicação estiver vazia, escreva: - Verificar compatibilidade com o fabricante
- NUNCA escreva o código do produto na APLICACAO.

========================================================
REGRAS DOS PARAGRAFOS DE TEXTO (aprenda com o exemplo):
- Parágrafo 1: explique o que é o produto, para que serve e qual problema resolve. 3 a 4 frases.
- Parágrafo 2: detalhe materiais, especificações técnicas, certificações e diferenciais. 3 a 4 frases.
- Ambos em terceira pessoa. Sem bullet points. Sem "nosso/nossa".
- NUNCA escreva parágrafos genéricos que sirvam para qualquer produto.

========================================================
AGORA CRIE O ANUNCIO PARA:

DADOS DO PRODUTO:
Produto: {product['produto']}
Marca: {product['marca']}
Codigo: {product['codigo']}
Aplicacao: {aplicacao if aplicacao else 'Não informada'}
Loja: {client_name}

FORMATO OBRIGATORIO — copie os rótulos EXATAMENTE como no exemplo:


[uma unica linha de titulo SEO]

DESCRICAO COMPLETA (Padrao Escalada Ecom)
--------------------------------------------------
O QUE VEM NA CAIXA
- 01 {product['produto']}
- Marca: {product['marca']}
- Codigo/Referencia: {product['codigo']}
--------------------------------------------------
APLICACAO
[bullets de aplicação/compatibilidade baseados nos dados fornecidos]
--------------------------------------------------
[Paragrafo 1 — específico ao produto, explica uso e benefício]

[Paragrafo 2 — materiais, especificações técnicas, diferenciais]
--------------------------------------------------
INSTITUCIONAL

A {client_name} oferece produtos de qualidade para abastecimento, manutenção e operação, disponibilizando soluções confiáveis para maior segurança e eficiência no dia a dia.

Nosso compromisso é entregar produtos de procedência, envio rápido e atendimento de confiança, ajudando sua operação a manter desempenho, organização e segurança."""

    MAX_TENTATIVAS = 8
    for tentativa in range(1, MAX_TENTATIVAS + 1):
        try:
            response = client.chat.completions.create(
                model="openai/gpt-oss-20b",
                messages=[
                    {"role": "system", "content": system_message},
                    {"role": "user", "content": prompt},
                ],
                temperature=0.5,
            )
            return response.choices[0].message.content
        except Exception as e:
            erro = str(e)
            if "429" in erro or "rate_limit" in erro.lower():
                match = re.search(r'try again in ([\d.]+)s', erro, re.IGNORECASE)
                espera = int(float(match.group(1))) + 5 if match else 65
                print(f"\n  ⏳ Rate limit (429) — aguardando {espera}s "
                      f"(tentativa {tentativa}/{MAX_TENTATIVAS})...")
                time.sleep(espera)
            else:
                raise
    raise Exception(f"Falhou após {MAX_TENTATIVAS} tentativas por rate limit.")


# ══════════════════════════════════════════════════════════
#  CRIAÇÃO DO GOOGLE DOC
# ══════════════════════════════════════════════════════════

def create_google_doc(drive, docs, title, content, folder_id):
    """Cria um Google Doc com o conteúdo gerado e salva na pasta TEXTOS."""
    # Cria o documento vazio na pasta certa
    meta = {
        "name": title,
        "mimeType": "application/vnd.google-apps.document",
        "parents": [folder_id],
    }
    doc = drive.files().create(body=meta, fields="id").execute()
    doc_id = doc["id"]

    # Insere o texto
    docs.documents().batchUpdate(
        documentId=doc_id,
        body={"requests": [{"insertText": {"location": {"index": 1}, "text": content}}]},
    ).execute()

    return doc_id


# ══════════════════════════════════════════════════════════
#  PROCESSAMENTO POR CLIENTE
# ══════════════════════════════════════════════════════════

def process_client(client_name, client_folder_id, drive, sheets, docs):
    """Processa todos os produtos de um cliente."""
    sep = "─" * 55
    print(f"\n{sep}")
    print(f"  Cliente: {client_name}")
    print(sep)

    # Planilha
    spreadsheet = find_main_spreadsheet(drive, client_folder_id)
    if not spreadsheet:
        print("  ⚠️  Nenhuma planilha encontrada. Pulando.")
        return 0, 0

    print(f"  📊 Planilha: {spreadsheet['name']}")

    # Lê produtos
    try:
        products = read_products_from_spreadsheet(drive, sheets, spreadsheet)
    except Exception as e:
        print(f"  ❌ Erro ao ler planilha: {e}")
        return 0, 0

    print(f"  📦 {len(products)} produtos encontrados")
    if not products:
        return 0, 0

    # Pasta TEXTOS
    textos_id = get_or_create_textos_folder(drive, client_folder_id)
    existing  = list_existing_docs(drive, textos_id)
    print(f"  📝 {len(existing)} docs já existentes em TEXTOS (serão pulados)")

    created = skipped = errors = 0

    for i, product in enumerate(products, 1):
        title = product["produto"]
        prefix = f"  [{i:>3}/{len(products)}]"

        if title in existing:
            print(f"{prefix} ⏭️  Já existe — {title[:45]}")
            skipped += 1
            continue

        print(f"{prefix} ✨ Gerando  — {title[:45]}...", end="", flush=True)
        try:
            content = generate_with_groq(product, client_name)
            create_google_doc(drive, docs, title, content, textos_id)
            print(" ✅")
            created += 1
            existing.add(title)           # marca como criado para esta sessão
            time.sleep(DELAY_ENTRE_PRODUTOS)
        except Exception as e:
            print(f" ❌ {e}")
            errors += 1

    print(f"\n  Resultado: {created} criados | {skipped} pulados | {errors} erros")
    return created, skipped


# ══════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════

def main():
    print("\n" + "═" * 55)
    print("  AUTOMAÇÃO MERCADO LIVRE — Geração de Anúncios")
    print("═" * 55)

    # Autenticação
    print("\n🔑 Autenticando com Google...")
    creds  = get_credentials()
    drive  = build("drive",  "v3", credentials=creds)
    sheets = build("sheets", "v4", credentials=creds)
    docs   = build("docs",   "v1", credentials=creds)
    print("✅ Autenticado!\n")

    # Pasta raiz — ID fixo
    print(f"📁 Usando pasta MERCADO LIVRE (ID: {PASTA_RAIZ_ID})")

    # Lista clientes
    clients = list_subfolders(drive, PASTA_RAIZ_ID)
    if not clients:
        print("⚠️  Nenhuma subpasta de cliente encontrada.")
        return

    print(f"\n👥 {len(clients)} clientes disponíveis:")
    for i, c in enumerate(clients, 1):
        print(f"   {i:>2}. {c['name']}")

    # Seleção
    print("\nQual cliente processar?")
    print("   0 = Todos os clientes")
    choice = input("Digite o número: ").strip()

    if choice == "0":
        selected = clients
    else:
        try:
            idx = int(choice) - 1
            selected = [clients[idx]]
        except (ValueError, IndexError):
            print("❌ Opção inválida.")
            return

    # Processa
    total_criados = total_pulados = 0
    for client in selected:
        c, s = process_client(client["name"], client["id"], drive, sheets, docs)
        total_criados  += c
        total_pulados  += s

    # Resumo final
    print("\n" + "═" * 55)
    print("  CONCLUÍDO!")
    print(f"  Documentos criados : {total_criados}")
    print(f"  Produtos pulados   : {total_pulados}")
    print("═" * 55 + "\n")


if __name__ == "__main__":
    main()