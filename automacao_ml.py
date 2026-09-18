#!/usr/bin/env python3
"""
============================================================
  AUTOMAÇÃO MERCADO LIVRE — Geração de Anúncios via OpenAI
============================================================
Lê produtos da planilha no Google Drive, gera título e
descrição via API da OpenAI e salva como Google Docs na
pasta TEXTOS do cliente.

SETUP RÁPIDO:
  1. pip install -r requirements.txt
  2. Baixe credentials.json do Google Cloud Console
     (veja INSTRUCOES.txt para o passo a passo)
  3. python automacao_ml.py
============================================================
"""

import base64
import datetime
import io
import os
import pickle
import re
import time
import unicodedata
from collections import defaultdict
from pathlib import Path

# Força o diretório de trabalho para a pasta onde este script está salvo
os.chdir(Path(__file__).resolve().parent)

import openpyxl
from openai import OpenAI, APIConnectionError, APITimeoutError
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# ─────────────────────────────────────────────────────────
#  CONFIGURAÇÕES — edite aqui se necessário
# ─────────────────────────────────────────────────────────

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "").strip()

# Modelo OpenAI a usar — pode sobrescrever via OPENAI_MODEL=... no .env
OPENAI_MODEL = os.getenv("OPENAI_MODEL", "gpt-5.2").strip()
OPENAI_RESEARCH_MODEL = os.getenv("OPENAI_RESEARCH_MODEL", OPENAI_MODEL).strip()
WEB_RESEARCH_ENABLED = os.getenv("WEB_RESEARCH_ENABLED", "1").strip().lower() not in (
    "0", "false", "nao", "n\u00e3o"
)

# ID da pasta MERCADO LIVRE no Google Drive (fixo — não precisa alterar)
PASTA_RAIZ_ID = "1H7r7kvGIuuqZByHaAVXxvkHA64852_if"

# Linha de início dos dados na planilha (pula cabeçalhos)
LINHA_INICIO = 3  # linha 3 = primeira linha de produto

# Palavras-chave para detecção automática de colunas (sem acentos — comparação normalizada)
_HEADER_KEYWORDS = {
    "produto":   ["produto", "product", "item", "descricao", "nome"],
    "marca":     ["marca", "brand", "fabricante"],
    "codigo":    ["codigo", "code", "ref", "referencia", "sku", "cod"],
    "aplicacao": ["aplicacao", "aplicavel", "compatib", "veiculo", "aplicac"],
}


def _norm(s):
    """Remove acentos e normaliza para comparação de cabeçalhos."""
    return ''.join(
        c for c in unicodedata.normalize('NFD', str(s).strip().lower())
        if unicodedata.category(c) != 'Mn'
    )
# Fallback caso nenhum cabeçalho seja encontrado
_COL_DEFAULTS = {"produto": 0, "marca": 1, "codigo": 2, "aplicacao": 3}

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


def get_sheet_names(drive, sheets_service, file_info):
    """Retorna lista de nomes de abas de uma planilha."""
    fid  = file_info["id"]
    mime = file_info["mimeType"]

    if mime == "application/vnd.google-apps.spreadsheet":
        meta = sheets_service.spreadsheets().get(
            spreadsheetId=fid, fields="sheets.properties.title"
        ).execute()
        return [s["properties"]["title"] for s in meta.get("sheets", [])]
    else:
        request = drive.files().get_media(fileId=fid)
        buf = io.BytesIO()
        dl = MediaIoBaseDownload(buf, request)
        done = False
        while not done:
            _, done = dl.next_chunk()
        buf.seek(0)
        wb = openpyxl.load_workbook(buf, read_only=True)
        names = list(wb.sheetnames)
        wb.close()
        return names


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
    """Retorna um set com os nomes dos arquivos já existentes em TEXTOS (com paginação)."""
    q = f"'{folder_id}' in parents and trashed=false"
    names = set()
    page_token = None
    while True:
        kwargs = dict(q=q, fields="nextPageToken, files(name)", pageSize=1000)
        if page_token:
            kwargs["pageToken"] = page_token
        r = drive.files().list(**kwargs).execute()
        for f in r.get("files", []):
            names.add(f["name"])
        page_token = r.get("nextPageToken")
        if not page_token:
            break
    return names


# ══════════════════════════════════════════════════════════
#  LEITURA DA PLANILHA
# ══════════════════════════════════════════════════════════

def detect_columns(header_rows):
    """Detecta índices de coluna a partir das linhas de cabeçalho da planilha.

    Varre as linhas de cabeçalho procurando palavras-chave conhecidas.
    Retorna um dict {campo: índice_coluna}. Usa _COL_DEFAULTS para campos
    não encontrados.
    """
    # best[field] = (prioridade_do_keyword, col_idx) — menor prioridade = mais específico
    best = {}
    for row in header_rows:
        for col_idx, cell in enumerate(row):
            cell_norm = _norm(cell)
            if not cell_norm or cell_norm in ("none", "nan"):
                continue
            for field, keywords in _HEADER_KEYWORDS.items():
                for priority, kw in enumerate(keywords):
                    if kw in cell_norm:
                        if field not in best or priority < best[field][0]:
                            best[field] = (priority, col_idx)
                        break  # usa o keyword de maior prioridade para esta célula

    mapping = {field: col_idx for field, (_, col_idx) in best.items()}

    # preenche campos não detectados com o fallback
    for field, default_idx in _COL_DEFAULTS.items():
        if field not in mapping:
            mapping[field] = default_idx

    return mapping


def _xlsx_cell(c):
    """Converte valor de célula xlsx para string, tratando datas formatadas como número."""
    if c is None:
        return ""
    if isinstance(c, (datetime.datetime, datetime.date)):
        from openpyxl.utils.datetime import to_excel
        dt = c if isinstance(c, datetime.datetime) else datetime.datetime(c.year, c.month, c.day)
        return str(int(to_excel(dt)))
    return str(c)


def _parse_rows(rows, col_map):
    """Converte linhas brutas em lista de produtos usando mapeamento dinâmico de colunas."""
    products = []
    for row in rows:
        if len(row) <= col_map["produto"]:
            continue

        def cell(idx):
            if idx is None or len(row) <= idx:
                return ""
            v = str(row[idx]).strip()
            return "" if v.lower() in ("none", "nan", "") else v

        produto   = cell(col_map["produto"])
        marca     = cell(col_map["marca"])
        aplicacao = cell(col_map["aplicacao"])
        codigo    = cell(col_map["codigo"]) or "SEM"

        if produto and produto.lower() not in ("produto", "none", "nan"):
            products.append({
                "produto":   produto,
                "marca":     marca,
                "aplicacao": aplicacao,
                "codigo":    codigo or "SEM",
            })
    return products


def read_products_from_spreadsheet(drive, sheets, file_info, sheets_filter=None):
    """Lê os produtos das abas da planilha (Google Sheets nativo ou .xlsx).

    sheets_filter: lista de nomes de abas a processar, ou None para todas.
    """
    fid   = file_info["id"]
    mime  = file_info["mimeType"]
    products = []

    if mime == "application/vnd.google-apps.spreadsheet":
        # ── Google Sheets nativo — busca todas as abas ────
        meta = sheets.spreadsheets().get(spreadsheetId=fid, fields="sheets.properties").execute()
        sheet_names = [s["properties"]["title"] for s in meta.get("sheets", [])]
        if sheets_filter:
            sheet_names = [s for s in sheet_names if s in sheets_filter]
        print(f"     📑 {len(sheet_names)} aba(s) selecionada(s): {', '.join(sheet_names)}")
        for sheet_name in sheet_names:
            # lê cabeçalhos (linhas 1 até LINHA_INICIO-1) para detectar colunas
            header_result = sheets.spreadsheets().values().get(
                spreadsheetId=fid,
                range=f"'{sheet_name}'!A1:Z{LINHA_INICIO - 1}"
            ).execute()
            col_map = detect_columns(header_result.get("values", []))
            print(f"       🗂  Colunas detectadas: produto={col_map['produto']} marca={col_map['marca']} "
                  f"codigo={col_map['codigo']} aplicacao={col_map['aplicacao']}")

            result = sheets.spreadsheets().values().get(
                spreadsheetId=fid,
                range=f"'{sheet_name}'!A{LINHA_INICIO}:Z2000"
            ).execute()
            rows = result.get("values", [])
            found = _parse_rows(rows, col_map)
            if found:
                print(f"       → Aba '{sheet_name}': {len(found)} produto(s)")
            products.extend(found)
    else:
        # ── Arquivo .xlsx — percorre todas as abas ────────
        request = drive.files().get_media(fileId=fid)
        buf = io.BytesIO()
        dl = MediaIoBaseDownload(buf, request)
        done = False
        while not done:
            _, done = dl.next_chunk()
        buf.seek(0)
        wb = openpyxl.load_workbook(buf, data_only=True)
        all_names = list(wb.sheetnames)
        filtered  = [s for s in all_names if s in sheets_filter] if sheets_filter else all_names
        print(f"     📑 {len(filtered)} aba(s) selecionada(s): {', '.join(filtered)}")
        for sheet_name in filtered:
            ws = wb[sheet_name]
            # lê cabeçalhos para detectar colunas
            header_rows = []
            for row in ws.iter_rows(min_row=1, max_row=LINHA_INICIO - 1, values_only=True):
                header_rows.append([str(c) if c is not None else "" for c in row])
            print(f"       🔍 DEBUG cabeçalhos lidos (linhas 1-{LINHA_INICIO-1}):")
            for i, hr in enumerate(header_rows):
                print(f"          Linha {i+1}: {hr[:10]}")
            col_map = detect_columns(header_rows)
            print(f"       🗂  Colunas detectadas: produto={col_map['produto']} marca={col_map['marca']} "
                  f"codigo={col_map['codigo']} aplicacao={col_map['aplicacao']}")

            rows = []
            for row in ws.iter_rows(min_row=LINHA_INICIO, values_only=True):
                rows.append([_xlsx_cell(c) for c in row])
            found = _parse_rows(rows, col_map)
            if found:
                print(f"       → Aba '{sheet_name}': {len(found)} produto(s)")
            products.extend(found)

    return products


# ══════════════════════════════════════════════════════════
#  GERAÇÃO VIA GEMINI
# ══════════════════════════════════════════════════════════

def _chamar_modelo(modelo, system_message, prompt):
    """Chama o modelo OpenAI."""
    if not OPENAI_API_KEY:
        raise Exception("OPENAI_API_KEY não configurada no .env")
    client = OpenAI(api_key=OPENAI_API_KEY)
    response = client.chat.completions.create(
        model=modelo,
        messages=[
            {"role": "system", "content": system_message},
            {"role": "user", "content": prompt},
        ],
        temperature=0.5,
    )
    return response.choices[0].message.content


def _pesquisar_aplicacao_completa(product):
    """Pesquisa compatibilidades e palavras-chave antes de gerar o anuncio final."""
    if not OPENAI_API_KEY:
        raise Exception("OPENAI_API_KEY nao configurada no .env")

    client = OpenAI(api_key=OPENAI_API_KEY)
    aplicacao_planilha = product.get("aplicacao", "") or "Nao informada"
    consulta = (
        f"produto: {product.get('produto', '')}; "
        f"marca: {product.get('marca', '')}; "
        f"codigo/referencia: {product.get('codigo', '')}; "
        f"aplicacao informada na planilha: {aplicacao_planilha}"
    )

    instructions = """Voce e um pesquisador tecnico de autopecas e SEO para comercio eletronico.
Pesquise na web a aplicacao completa do produto informado.

REGRAS:
- Priorize catalogos de fabricantes, distribuidores, catalogos de autopecas e fontes tecnicas confiaveis.
- Use o codigo/referencia como identificador principal e compare marca, produto, modelo, motor, versao e anos.
- Separe claramente o que foi confirmado do que e incerto ou conflitante.
- Nao invente compatibilidades e nao trate resultado de busca generico como confirmacao.
- Retorne uma lista organizada por montadora, modelo, motor, versao e anos quando esses dados estiverem confirmados.
- Retorne uma secao separada com palavras-chave relevantes: nome principal, sinonimos, nomes usados no mercado, codigos, modelos compativeis e termos de alta intencao de compra.
- Use termos encontrados em catalogos, anuncios e fontes tecnicas; nao invente volume de busca nem use palavras-chave sem relacao comprovada com o produto.
- Inclua as URLs das fontes usadas para cada grupo de compatibilidade.
- Se nao encontrar confirmacao suficiente, diga isso explicitamente e preserve a aplicacao da planilha como referencia nao verificada.
"""

    response = client.responses.create(
        model=OPENAI_RESEARCH_MODEL,
        tools=[{"type": "web_search"}],
        instructions=instructions,
        input=f"Pesquise a aplicacao completa para:\n{consulta}",
    )
    return (response.output_text or "").strip()


def generate_with_groq(product, client_name):
    """Gera o texto via IA com rotação automática de modelos (Gemini + Groq)."""

    system_message = """Você é um especialista em criação de descrições profissionais para anúncios de autopeças no Mercado Livre.
Sua função é criar uma descrição completa, clara, comercial, segura e otimizada para conversão, baseada no seu conhecimento técnico sobre o produto.

REGRAS ABSOLUTAS — NUNCA QUEBRE:
1. Responda SOMENTE com o conteúdo do anúncio. Zero introduções, zero "Aqui está", zero comentários.
2. NUNCA use asteriscos (*), hashtags (#), aspas ou qualquer markdown.
3. Não invente informações técnicas, aplicações, anos, medidas, códigos, compatibilidades ou palavras-chave. Para aplicações e palavras-chave, use somente a planilha e a pesquisa fornecida.
4. Não diga que o produto é original se isso não estiver confirmado.
5. Use linguagem profissional, simples e confiável.
6. Não use emojis.
7. Sempre oriente o comprador a confirmar a compatibilidade antes da compra.
8. Use títulos em letras maiúsculas.
9. Use listas em tópicos com hífen.
10. Não use tabelas.
11. A descrição deve ser completa, mas objetiva e direta."""

    aplicacao = product['aplicacao'] if product['aplicacao'] else ''
    pesquisa_aplicacao = ''
    if WEB_RESEARCH_ENABLED:
        try:
            print("     Pesquisando a aplicacao completa na web...")
            pesquisa_aplicacao = _pesquisar_aplicacao_completa(product)
            if pesquisa_aplicacao:
                print("     Pesquisa de aplicacao concluida.")
            else:
                print("     Pesquisa nao retornou dados; usando a planilha.")
        except Exception as e:
            print(f"     Aviso: pesquisa de aplicacao indisponivel ({str(e)[:160]}).")

    pesquisa_aplicacao = pesquisa_aplicacao or "Pesquisa indisponivel. Nao amplie a aplicacao alem do que consta na planilha."

    prompt = f"""Crie um anúncio completo para o Mercado Livre com base nos dados abaixo.
Use seu conhecimento técnico sobre o produto para complementar as informações fornecidas.

DADOS DO PRODUTO:
Produto: {product['produto']}
Marca: {product['marca']}
Codigo: {product['codigo']}
Aplicacao: {aplicacao if aplicacao else 'Não informada'}

========================================================
PESQUISA DE APLICACAO COMPLETA:
{pesquisa_aplicacao}

========================================================
REGRAS DOS TITULOS SEO:
- Gere dois títulos diferentes usando as palavras-chave pesquisadas, sem keyword stuffing.
- TITULO CURTO: máximo de 60 caracteres, para o título principal do Mercado Livre. Seja direto e priorize produto, função e modelos compatíveis confirmados. Sem a marca e sem o código.
- TITULO LONGO: máximo de 200 caracteres, mais completo e natural, podendo incluir marca, código, modelos, motores, anos e sinônimos confirmados.
- Conte os caracteres antes de responder e nunca ultrapasse os limites.
- Não repita o mesmo título nos dois campos.
- RUIM: "FILTRO COMBUSTIVEL"
- BOM: "Filtro Combustivel Gol Polo Civic HB20 Clio 1.0 1.4 1.6"

REGRAS DA SECAO APLICACAO DO PRODUTO:
- Use a aplicação da planilha como base e complemente somente com compatibilidades confirmadas na pesquisa acima.
- NUNCA transforme uma possibilidade, conflito ou resultado genérico em compatibilidade confirmada.
- Agrupe por montadora com tópicos. Inclua modelo, motor e anos quando disponíveis.
- Ao final da seção, inclua SEMPRE esta observação:
  A aplicação pode variar conforme versão, ano, motor ou configuração do veículo. Antes da compra, confira o código da peça antiga, as medidas e as fotos do anúncio.

REGRAS DAS PERGUNTAS FREQUENTES:
- Crie de 6 a 10 perguntas e respostas objetivas baseadas no produto.
- Responda sempre com segurança, sem inventar informações.
- Inclua perguntas como: serve no meu veículo, qual a marca, qual o código, o produto é novo, o que acompanha, precisa de mecânico, como confirmar compatibilidade.

========================================================
FORMATO OBRIGATORIO — copie os rótulos EXATAMENTE:

TITULO CURTO (ATÉ 60 CARACTERES)

[uma unica linha com no maximo 60 caracteres]

TITULO LONGO (ATÉ 200 CARACTERES)

[uma unica linha com no maximo 200 caracteres]

DESCRICAO COMPLETA
--------------------------------------------------
O QUE VAI NA CAIXA

- 01 {product['produto']}
[se for kit, liste os itens separadamente com base no seu conhecimento técnico]
--------------------------------------------------
APLICACAO DO PRODUTO

{('[aplicacao dos veiculos agrupada por montadora em topicos]') if not aplicacao else aplicacao}

A aplicacao pode variar conforme versao, ano, motor ou configuracao do veiculo. Antes da compra, confira o codigo da peca antiga, as medidas e as fotos do anuncio.
--------------------------------------------------
DESCRICAO DO PRODUTO

[abertura comercial curta: o que e o produto, para qual aplicacao e indicado e qual problema ele resolve. Tom direto e confiavel. 2 a 3 frases.]
--------------------------------------------------
CARACTERISTICAS PRINCIPAIS

- Produto: {product['produto']}
- Marca: {product['marca']}
- Codigo/Referencia: {product['codigo']}
[adicione: material, medidas, lado/posicao, condicao (novo), funcao da peca e especificacoes tecnicas relevantes que voce conhecer]
--------------------------------------------------
FUNCAO DA PECA

[explique de forma simples a funcao do produto no veiculo. Adapte ao tipo de peca: filtro, kit transmissao, sensor, bomba, suspensao, peca eletrica etc. 2 a 3 frases.]
--------------------------------------------------
BENEFICIOS

- Ideal para reposicao
- Auxilia na manutencao correta do veiculo
- Ajuda a evitar falhas causadas por peca desgastada
- Boa opcao para manutencao preventiva ou corretiva
- Indicado para uso diario ou profissional
- Bom custo-beneficio
- Recomendado para oficinas, mecanicos e proprietarios
--------------------------------------------------
CUIDADOS IMPORTANTES

- Confira a compatibilidade antes da compra
- Verifique o codigo da peca antiga
- Compare as fotos do anuncio com a peca instalada no veiculo
- Confirme ano, modelo, motor e versao
- Nao force a instalacao
- Em caso de duvida, envie uma pergunta antes da compra
--------------------------------------------------
ENVIO E PRAZO DE ENTREGA

O envio e realizado pelo Mercado Livre, conforme as opcoes disponiveis no momento da compra.
O prazo de entrega e calculado automaticamente pela plataforma de acordo com o CEP informado pelo comprador.
Apos a confirmacao do pagamento, o pedido sera separado e enviado com agilidade.
--------------------------------------------------
PERGUNTAS FREQUENTES

[crie de 6 a 10 perguntas e respostas objetivas sobre este produto especifico]
--------------------------------------------------
Ainda esta com duvida sobre a compatibilidade? Envie sua pergunta antes da compra. Nossa equipe esta a disposicao para ajudar voce a escolher a peca correta para o seu veiculo."""

    MAX_TENTATIVAS = 4
    for tentativa in range(1, MAX_TENTATIVAS + 1):
        try:
            return _chamar_modelo(OPENAI_MODEL, system_message, prompt)
        except (APIConnectionError, APITimeoutError) as e:
            espera = 15
            print(f"\n  ⏳ Erro de conexão ({str(e)[:50]}...) — aguardando {espera}s "
                  f"(tentativa {tentativa}/{MAX_TENTATIVAS})...")
            time.sleep(espera)
        except Exception as e:
            erro = str(e)
            if "429" in erro or "rate_limit" in erro.lower():
                match = re.search(r'try again in ([\d.]+)s', erro, re.IGNORECASE)
                espera = int(float(match.group(1))) + 5 if match else 60
                print(f"\n  ⏳ Rate limit — aguardando {espera}s "
                      f"(tentativa {tentativa}/{MAX_TENTATIVAS})...")
                time.sleep(espera)
            elif any(k in erro.lower() for k in ("connect", "timeout", "read", "conexão", "interrompida", "connection")):
                espera = 15
                print(f"\n  ⏳ Erro de conexão ({erro[:50]}...) — aguardando {espera}s "
                      f"(tentativa {tentativa}/{MAX_TENTATIVAS})...")
                time.sleep(espera)
            else:
                print(f"\n  ❌ Erro inesperado: {erro}")
                raise
    raise Exception(f"Falhou após {MAX_TENTATIVAS} tentativas.")


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

    # Garante que haja texto para inserir, evitando erro 400 da API do Docs
    if not content or not str(content).strip():
        content = "[ERRO: A inteligência artificial não retornou nenhum texto para este anúncio. Verifique o prompt ou os dados do produto.]"

    # Insere o texto
    docs.documents().batchUpdate(
        documentId=doc_id,
        body={"requests": [{"insertText": {"location": {"index": 1}, "text": content}}]},
    ).execute()

    return doc_id


# ══════════════════════════════════════════════════════════
#  PROCESSAMENTO POR CLIENTE
# ══════════════════════════════════════════════════════════

def process_client(client_name, client_folder_id, drive, sheets, docs, sheets_filter=None):
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
        products = read_products_from_spreadsheet(drive, sheets, spreadsheet, sheets_filter=sheets_filter)
    except Exception as e:
        print(f"  ❌ Erro ao ler planilha: {e}")
        return 0, 0

    print(f"  📦 {len(products)} produtos encontrados")
    if not products:
        return 0, 0

    # Mostra os primeiros nomes lidos da planilha para conferência
    for p in products[:5]:
        print(f"       → {p['produto'][:60]}")
    if len(products) > 5:
        print(f"       ... e mais {len(products) - 5}")

    # Pasta TEXTOS
    textos_id = get_or_create_textos_folder(drive, client_folder_id)
    existing  = list_existing_docs(drive, textos_id)
    print(f"  📝 {len(existing)} docs já existentes em TEXTOS (serão pulados)")

    # Se houver docs existentes, mostra os primeiros para diagnóstico
    if existing:
        sample = sorted(existing)[:5]
        for name in sample:
            print(f"       já existe: {name[:60]}")
        if len(existing) > 5:
            print(f"       ... e mais {len(existing) - 5}")

    created = skipped = errors = 0

    # Rastreia quantas vezes cada nome de produto aparece para gerar títulos únicos
    name_seen: dict = defaultdict(int)

    for i, product in enumerate(products, 1):
        base_title = f"{product['produto']} - {product['codigo']}"
        name_seen[base_title] += 1
        occurrence = name_seen[base_title]
        title = base_title if occurrence == 1 else f"{base_title} ({occurrence})"

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

    print(f"\n🤖 Modelo de IA: {OPENAI_MODEL}  [OpenAI]")

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
