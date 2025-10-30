# -*- coding: utf-8 -*-
"""
core_xylella.py — Cloud/Streamlit (OCR Azure + Parser Colab + Writer por requisição)

API esperada pela UI (xylella_processor.py):
    • process_pdf_sync(pdf_path) -> List[List[Dict]]   # devolve listas de amostras por requisição
    • write_to_template(rows_per_req, out_base_path, expected_count=None, source_pdf=None) -> List[str]

Requer:
  - AZURE_API_KEY, AZURE_ENDPOINT (env)
  - TEMPLATE_PATH (env) ou ficheiro 'TEMPLATE_PXf_SGSLABIP1056.xlsx' ao lado do core
  - OUTPUT_DIR (env) — diretório onde guardar .xlsx e _ocr_debug.txt
"""

from __future__ import annotations
import os, re, io, time, json, shutil, requests
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Any, Tuple

from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment

# ───────────────────────────────────────────────
# Caminhos / Ambiente
# ───────────────────────────────────────────────
BASE_DIR = Path(__file__).parent
OUTPUT_DIR = Path(os.environ.get("OUTPUT_DIR", BASE_DIR / "Output"))
OUTPUT_DIR.mkdir(exist_ok=True)

TEMPLATE_PATH = Path(os.environ.get("TEMPLATE_PATH", BASE_DIR / "TEMPLATE_PXf_SGSLABIP1056.xlsx"))
if not TEMPLATE_PATH.exists():
    # não falha já — o app.py normalmente garante isto; fazemos só aviso
    print(f"ℹ️ Aviso: TEMPLATE não encontrado em {TEMPLATE_PATH}. Será verificado no momento da escrita.")

AZURE_API_KEY = os.environ.get("AZURE_API_KEY", "")
AZURE_ENDPOINT = os.environ.get("AZURE_ENDPOINT", "")
MODEL_ID = os.environ.get("AZURE_MODEL_ID", "prebuilt-document")

# ───────────────────────────────────────────────
# Estilos Excel
# ───────────────────────────────────────────────
GREEN = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
RED   = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
GRAY  = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
BOLD  = Font(bold=True, color="000000")
ITALIC= Font(italic=True, color="555555")

# ───────────────────────────────────────────────
# Utilitários
# ───────────────────────────────────────────────
def _is_valid_date(v) -> bool:
    try:
        datetime.strptime(str(v).strip(), "%d/%m/%Y")
        return True
    except Exception:
        return False

def _to_dt(v):
    try:
        return datetime.strptime(str(v).strip(), "%d/%m/%Y")
    except Exception:
        return v

def extract_all_text(result_json: Dict[str, Any]) -> str:
    lines = []
    for pg in result_json.get("analyzeResult", {}).get("pages", []):
        for ln in pg.get("lines", []):
            txt = (ln.get("content") or ln.get("text") or "").strip()
            if txt:
                lines.append(txt)
    return "\n".join(lines)

# ───────────────────────────────────────────────
# OCR Azure (PDF direto)
# ───────────────────────────────────────────────
def azure_analyze_pdf(pdf_path: str) -> Dict[str, Any]:
    if not AZURE_API_KEY or not AZURE_ENDPOINT:
        raise RuntimeError("Azure não configurado (AZURE_API_KEY/AZURE_ENDPOINT).")

    url = f"{AZURE_ENDPOINT.rstrip('/')}/formrecognizer/documentModels/{MODEL_ID}:analyze?api-version=2023-07-31"
    headers = {"Ocp-Apim-Subscription-Key": AZURE_API_KEY, "Content-Type": "application/pdf"}

    with open(pdf_path, "rb") as f:
        resp = requests.post(url, data=f.read(), headers=headers, timeout=90)
    if resp.status_code != 202:
        raise RuntimeError(f"Azure analyze falhou: {resp.status_code} {resp.text}")

    op = resp.headers.get("Operation-Location")
    if not op:
        raise RuntimeError("Azure não devolveu Operation-Location.")

    for _ in range(60):
        time.sleep(1.2)
        r = requests.get(op, headers={"Ocp-Apim-Subscription-Key": AZURE_API_KEY}, timeout=30)
        j = r.json()
        if j.get("status") == "succeeded":
            return j
        if j.get("status") == "failed":
            raise RuntimeError(f"OCR Azure falhou: {j}")
    raise RuntimeError("Timeout a aguardar OCR Azure.")

# ───────────────────────────────────────────────
# Parser — Secções do Colab (adaptadas)
# ───────────────────────────────────────────────
NATUREZA_KEYWORDS = [
    "ramos","folhas","ramosefolhas","ramosc/folhas","material","materialherbalho",
    "materialherbário","materialherbalo","natureza","insetos","sementes","solo"
]
TIPO_RE = re.compile(r"\b(Simples|Composta|Individual)\b", re.I)

def _looks_like_natureza(txt: str) -> bool:
    t = re.sub(r"\s+", "", (txt or "").lower())
    return any(k in t for k in NATUREZA_KEYWORDS)

def _clean_ref(raw: str) -> str:
    s = (raw or "").strip()
    s = re.sub(r"\s*/\s*", "/", s)
    s = re.sub(r"/{2,}", "/", s)
    s = re.sub(r"[A-Za-z]+", lambda m: m.group(0).upper(), s)
    s = s.replace("LUT", "LVT")
    s = re.sub(r"\s+", "", s)
    s = re.sub(r"[^A-Z0-9/]+$", "", s)
    return s

def detect_requisicoes(full_text: str):
    """Conta quantas requisições DGAV→SGS existem no texto OCR de um PDF."""
    pattern = re.compile(
        r"PROGRAMA\s+NACIONAL\s+DE\s+PROSPE[ÇC][AÃ]O\s+DE\s+PRAGAS\s+DE\s+QUARENTENA",
        re.IGNORECASE,
    )
    matches = list(pattern.finditer(full_text))
    count = len(matches)
    positions = [m.start() for m in matches]
    if count == 0:
        print("🔍 Nenhum cabeçalho encontrado — assumido 1 requisição.")
        count = 1
    else:
        print(f"🔍 Detetadas {count} requisições no ficheiro (posições: {positions})")
    return count, positions


def split_if_multiple_requisicoes(full_text: str):
    """
    Divide o texto OCR em blocos distintos (requisições DGAV→SGS),
    tolerando erros e quebras do OCR.
    """
    # Normaliza o texto (remove múltiplos espaços e capitaliza)
    text = re.sub(r"\s+", " ", full_text, flags=re.M).upper()

    # Padrão flexível: aceita variações e pequenas falhas no OCR
    pattern = re.compile(
        r"P\s*R\s*O\s*G\s*R\s*A\s*M\s*A\s+N\s*A\s*C\s*I\s*O\s*N\s*A\s*L\s+DE\s+PROSPE[CÇ]\s*AO\s+DE\s+PRAGAS\s+DE\s+QUARENTENA",
        re.IGNORECASE,
    )

    # Encontrar todas as ocorrências
    marks = [m.start() for m in pattern.finditer(text)]
    if not marks:
        print("🔍 Nenhum cabeçalho encontrado — tratado como 1 requisição.")
        return [full_text]
    if len(marks) == 1:
        print("🔍 Apenas 1 cabeçalho — 1 requisição detectada.")
        return [full_text]

    # Adiciona o fim do texto como limite final
    marks.append(len(text))
    blocos = []

    for i in range(len(marks) - 1):
        start = max(0, marks[i] - 200)           # inclui parte do cabeçalho anterior
        end = min(len(text), marks[i + 1] + 200)
        bloco = text[start:end].strip()
        if len(bloco) > 300:
            blocos.append(bloco)
        else:
            print(f"⚠️ Bloco {i+1} demasiado pequeno ({len(bloco)} chars).")

    print(f"🔍 Detetadas {len(blocos)} requisições distintas (por cabeçalho).")
    return blocos


def extract_context_from_text(full_text: str):
    """Extrai informações gerais da requisição (zona, DGAV, datas, nº de amostras)."""
    ctx = {}

    # Zona
    m_zona = re.search(r"Xylella\s+fastidiosa\s*\(([^)]+)\)", full_text, re.I)
    ctx["zona"] = m_zona.group(1).strip() if m_zona else "Zona Isenta"

    # DGAV / Responsável
    responsavel, dgav = None, None
    m_hdr = re.search(
        r"Amostra(?:s|\(s\))?\s*colhida(?:s|\(s\))?\s*por\s*DGAV\s*[:\-]?\s*(.*)",
        full_text, re.IGNORECASE,
    )
    if m_hdr:
        tail = full_text[m_hdr.end():]
        linhas = [m_hdr.group(1)] + tail.splitlines()
        for ln in linhas[:4]:
            ln = (ln or "").strip()
            if ln:
                responsavel = ln
                break
        if responsavel:
            responsavel = re.sub(r"\S+@dgav\.pt|\S+@\S+", "", responsavel, flags=re.I)
            responsavel = re.sub(r"PROGRAMA.*|Data.*|N[º°].*", "", responsavel, flags=re.I)
            responsavel = re.sub(r"[:;,.\-–—]+$", "", responsavel).strip()

    if responsavel:
        dgav = f"DGAV {responsavel}".strip() if not re.match(r"^DGAV\b", responsavel, re.I) else responsavel
    else:
        m_d = re.search(r"\bDGAV(?:\s+[A-Za-zÀ-ÿ?]+){1,4}", full_text)
        dgav = re.sub(r"[:;,.\-–—]+$", "", m_d.group(0)).strip() if m_d else None

    ctx["dgav"] = dgav
    ctx["responsavel_colheita"] = None

    # Datas de colheita
    colheita_map = {}
    for m in re.finditer(r"(\d{1,2}/\d{1,2}/\d{4})\s*\(\s*(\*+)\s*\)", full_text):
        colheita_map[f"({m.group(2).replace(' ', '')})"] = m.group(1)
    if not colheita_map:
        m_simple = re.search(r"Data\s+de\s+colheita\s*[:\-\s]*([0-9/\-\s]+)", full_text, re.I)
        if m_simple:
            only_date = re.sub(r"\s+", "", m_simple.group(1))
            for key in ("(*)", "(**)", "(***)"):
                colheita_map[key] = only_date
    default_colheita = next(iter(colheita_map.values()), "")
    ctx["colheita_map"] = colheita_map
    ctx["default_colheita"] = default_colheita

    # Data de envio
    m_envio = re.search(
        r"Data\s+(?:do|de)\s+envio(?:\s+ao\s+laborat[oó]rio)?[:\-\s]*([0-9/\-\s]+)",
        full_text, re.I,
    )
    if m_envio:
        ctx["data_envio"] = re.sub(r"\s+", "", m_envio.group(1))
    elif default_colheita:
        ctx["data_envio"] = default_colheita
    else:
        ctx["data_envio"] = datetime.now().strftime("%d/%m/%Y")

    # Nº de amostras declaradas
    flat = re.sub(r"\s+", " ", full_text)
    m_decl = re.search(r"N[º°]?\s*de\s*amostras(?:\s+neste\s+envio)?\s*[:\-]?\s*(\d{1,4})", flat, re.I)
    ctx["declared_samples"] = int(m_decl.group(1)) if m_decl else None

    return ctx

def parse_xylella_tables(result_json, context, req_id=None):
    """Extrai as amostras das tabelas Azure OCR, aplicando o contexto da requisição."""
    out = []
    tables = result_json.get("analyzeResult", {}).get("tables", [])
    if not tables:
        print("⚠️ Nenhuma tabela encontrada.")
        return out

    for t in tables:
        nc = max(c.get("columnIndex", 0) for c in t.get("cells", [])) + 1
        nr = max(c.get("rowIndex", 0) for c in t.get("cells", [])) + 1
        grid = [[""] * nc for _ in range(nr)]
        for c in t.get("cells", []):
            grid[c["rowIndex"]][c["columnIndex"]] = clean_value(c.get("content", ""))

        for row in grid:
            if not row or not any(row):
                continue
            ref = _clean_ref(row[0]) if len(row) > 0 else ""
            if not ref or re.match(r"^\D+$", ref):
                continue

            hospedeiro = row[2] if len(row) > 2 else ""
            obs        = row[3] if len(row) > 3 else ""

            if _looks_like_natureza(hospedeiro):
                hospedeiro = ""

            tipo = ""
            joined = " ".join([x for x in row if isinstance(x, str)])
            m_tipo = re.search(r"\b(Simples|Composta|Individual)\b", joined, re.I)
            if m_tipo:
                tipo = m_tipo.group(1).capitalize()
                obs = re.sub(r"\b(Simples|Composta|Individual)\b", "", obs, flags=re.I).strip()

            datacolheita = context.get("default_colheita", "")
            m_ast = re.search(r"\(\s*\*+\s*\)", joined)
            if m_ast:
                mark = re.sub(r"\s+", "", m_ast.group(0))
                datacolheita = context.get("colheita_map", {}).get(mark, datacolheita)

            if obs.strip().lower() in ("simples", "composta", "individual"):
                obs = ""

            out.append({
                "requisicao_id": req_id,
                "datarececao": context["data_envio"],
                "datacolheita": datacolheita,
                "referencia": ref,
                "hospedeiro": hospedeiro,
                "tipo": tipo,
                "zona": context["zona"],
                "responsavelamostra": context["dgav"],
                "responsavelcolheita": context["responsavel_colheita"],
                "observacoes": obs.strip(),
                "procedure": "XYLELLA",
                "datarequerido": context["data_envio"],
                "Score": ""
            })

    print(f"✅ {len(out)} amostras extraídas no total (req_id={req_id}).")
    return out
    
def parse_xylella_from_result(result_json, pdf_name, txt_path=None):
    """
    Analisa o resultado OCR e extrai todas as requisições de um PDF Xylella.
    Divide o documento em blocos e associa as tabelas a cada bloco.
    """
    import os

    print(f"📄 A processar: {pdf_name}")

    # 1️⃣ Ler texto global do OCR
    if txt_path and os.path.exists(txt_path):
        with open(txt_path, "r", encoding="utf-8") as f:
            full_text = f.read()
        print(f"📝 Contexto extraído de {os.path.basename(txt_path)}")
    else:
        print("⚠️ Ficheiro de texto global não encontrado — fallback: 1.ª página.")
        first_page_text = "\n".join(
            line.get("content", "")
            for line in result_json.get("analyzeResult", {}).get("pages", [])[0].get("lines", [])
        )
        full_text = first_page_text

    # 2️⃣ Detetar requisições
    count, _ = detect_requisicoes(full_text)

    # 3️⃣ Documento com apenas uma requisição
    if count <= 1:
        print("📄 Documento contém apenas uma requisição.")
        context = extract_context_from_text(full_text)
        amostras = parse_xylella_tables(result_json, context, req_id=1)

        if not amostras:
            append_process_log(pdf_name, 1, 0, context.get("declared_samples"),
                               out_path=None, status="Vazia",
                               error_msg="Sem amostras válidas.")
            return []

        out_path = write_to_template(amostras, os.path.basename(pdf_name),
                                     expected_count=context.get("declared_samples"),
                                     source_pdf=pdf_name)
        append_process_log(pdf_name, 1, len(amostras),
                           context.get("declared_samples"), out_path=out_path, status="OK")
        return amostras

    # 4️⃣ Documento com múltiplas requisições
    blocos = split_if_multiple_requisicoes(full_text)
    print(f"📄 Documento dividido em {len(blocos)} requisições.")

    all_samples = []
    all_tables = result_json.get("analyzeResult", {}).get("tables", [])

    for i, bloco in enumerate(blocos, start=1):
        print(f"\n🔹 A processar requisição {i}/{len(blocos)}...")

        try:
            bloco = re.sub(r"[ \t]+", " ", bloco.replace("\r", ""))
            context = extract_context_from_text(bloco)

            refs_bloco = re.findall(
                r"\b\d{1,3}/[A-Z]{0,2}/DGAV(?:-[A-Z0-9/]+)?|\b\d{2,4}/\d{2,4}/[A-Z0-9\-]+",
                bloco, re.I
            )
            refs_bloco = [r.strip() for r in refs_bloco if len(r.strip()) > 4]
            print(f"   ↳ {len(refs_bloco)} referências detetadas no bloco {i}")

            tables_filtradas = [
                t for t in all_tables
                if any(ref in " ".join(c.get("content", "") for c in t.get("cells", []))
                       for ref in refs_bloco)
            ] or all_tables

            result_local = {"analyzeResult": {"tables": tables_filtradas}}
            amostras = parse_xylella_tables(result_local, context, req_id=i)

            if not amostras:
                append_process_log(pdf_name, i, 0, context.get("declared_samples"),
                                   out_path=None, status="Vazia",
                                   error_msg="Sem amostras válidas (OCR incompleto).")
                continue

            base = os.path.splitext(os.path.basename(pdf_name))[0]
            out_name = f"{base}_req{i}.xlsx"
            out_path = write_to_template(amostras, out_name,
                                         expected_count=context.get("declared_samples"),
                                         source_pdf=pdf_name)
            append_process_log(pdf_name, i, len(amostras),
                               context.get("declared_samples"), out_path=out_path, status="OK")
            print(f"✅ Requisição {i} gravada em: {out_path}")
            all_samples.extend(amostras)

        except Exception as e:
            print(f"❌ Erro na requisição {i}: {e}")
            append_process_log(pdf_name, i, 0, None,
                               out_path=None, status="Erro", error_msg=str(e))

    print(f"\n🏁 Concluído: {len(blocos)} requisições processadas, {len(all_samples)} amostras no total.")
    return all_samples
    
# ───────────────────────────────────────────────
# Parser: dividir e extrair requisições
# ───────────────────────────────────────────────
def parse_all_requisitions(result_json: Dict[str, Any], pdf_name: str, txt_path: str | None) -> List[List[Dict[str, Any]]]:
    """Divide o documento em blocos (requisições) e extrai amostras por bloco."""
    # Texto global do OCR
    if txt_path and os.path.exists(txt_path):
        full_text = Path(txt_path).read_text(encoding="utf-8")
        print(f"📝 Contexto extraído de {os.path.basename(txt_path)}")
    else:
        full_text = extract_all_text(result_json)
        print("⚠️ Ficheiro OCR não encontrado — a usar texto direto do OCR.")

    # Deteção de requisições (cabeçalhos)
    count, _ = detect_requisicoes(full_text)
    if count == 0:
        print("⚠️ Nenhum cabeçalho detectado — assumido 1 requisição.")
        count = 1

    all_tables = result_json.get("analyzeResult", {}).get("tables", []) or []
    out: List[List[Dict[str, Any]]] = []

    # Documento simples (1 requisição)
    if count <= 1:
        context = extract_context_from_text(full_text)
        amostras = parse_xylella_tables(result_json, context, req_id=1)
        return [amostras] if amostras else []

    # Documento com múltiplas requisições
    blocos = split_if_multiple_requisicoes(full_text)
    print(f"📄 Documento dividido em {len(blocos)} requisições distintas.")

    for i, bloco in enumerate(blocos, start=1):
        try:
            context = extract_context_from_text(bloco)
            refs_bloco = re.findall(r"\b\d{7,8}\b|\b\d{2,4}/\d{2,4}/[A-Z0-9\-]+\b", bloco, re.I)

            # Filtra tabelas correspondentes a esta requisição
            tables_filtradas = []
            for t in all_tables:
                joined = " ".join(c.get("content", "") for c in t.get("cells", []))
                if any(ref in joined for ref in refs_bloco):
                    tables_filtradas.append(t)

            if not tables_filtradas and i == 1:
                print("⚠️ Nenhuma tabela filtrada — usar todas por segurança.")
                tables_filtradas = all_tables

            local = {"analyzeResult": {"tables": tables_filtradas}}
            amostras = parse_xylella_tables(local, context, req_id=i)

            if amostras:
                print(f"✅ Requisição {i}: {len(amostras)} amostras.")
                out.append(amostras)
            else:
                print(f"⚠️ Requisição {i} sem amostras extraídas.")

        except Exception as e:
            print(f"❌ Erro na requisição {i}: {e}")

    return out

# ───────────────────────────────────────────────
# Escrita no TEMPLATE — 1 ficheiro por requisição
# ───────────────────────────────────────────────

def to_datetime(value: str):
    """Converte string dd/mm/yyyy para objeto datetime (ou devolve None se inválida)."""
    try:
        return datetime.strptime(str(value).strip(), "%d/%m/%Y")
    except Exception:
        return None
        
def write_to_template(ocr_rows, out_name, expected_count=None, source_pdf=None):
    """
    Escreve as linhas extraídas no template base.
    Cria um ficheiro Excel no diretório OUTPUT_DIR com o nome indicado.
    Campo de observações (coluna I) é sempre deixado vazio.
    Inclui:
      • Validação do nº de amostras (E1:F1)
      • Origem real do PDF (G1:J1)
      • Data/hora de processamento (K1:L1)
      • Conversão automática de datas
      • Validação de campos obrigatórios
      • Fórmula Data requerido = Data Receção + 30 dias
    """
    from openpyxl import load_workbook
    from openpyxl.styles import PatternFill, Font, Alignment
    from datetime import datetime
    import os

    if not ocr_rows:
        print(f"⚠️ {out_name}: sem linhas para escrever.")
        return None

    if not os.path.exists(TEMPLATE_PATH):
        raise FileNotFoundError(f"Template não encontrado: {TEMPLATE_PATH}")

    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.worksheets[0]
    start_row = 4

    # 🎨 Estilos
    yellow_fill = PatternFill(start_color="FFFACD", end_color="FFFACD", fill_type="solid")
    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    gray_fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
    bold_center = Font(bold=True, color="000000")

    def is_valid_date(value: str) -> bool:
        try:
            if isinstance(value, datetime):
                return True
            datetime.strptime(str(value).strip(), "%d/%m/%Y")
            return True
        except Exception:
            return False

    def to_datetime(value: str):
        try:
            return datetime.strptime(str(value).strip(), "%d/%m/%Y")
        except Exception:
            return None

    # 🧹 Limpar linhas anteriores (mantém apenas cabeçalhos)
    for row in range(start_row, 201):
        for col in range(1, 13):
            ws.cell(row=row, column=col).value = None
            ws.cell(row=row, column=col).fill = PatternFill(fill_type=None)

    # ✍️ Escrever novas linhas
    for idx, row in enumerate(ocr_rows, start=start_row):
        rececao_val = row.get("datarececao", "")
        colheita_val = row.get("datacolheita", "")

        cell_A = ws[f"A{idx}"]
        cell_B = ws[f"B{idx}"]

        if is_valid_date(rececao_val):
            cell_A.value = to_datetime(rececao_val)
        else:
            cell_A.value = rececao_val
            cell_A.fill = red_fill

        if is_valid_date(colheita_val):
            cell_B.value = to_datetime(colheita_val)
        else:
            cell_B.value = colheita_val
            cell_B.fill = red_fill

        ws[f"C{idx}"] = row.get("referencia", "")
        ws[f"D{idx}"] = row.get("hospedeiro", "")
        ws[f"E{idx}"] = row.get("tipo", "")
        ws[f"F{idx}"] = row.get("zona", "")
        ws[f"G{idx}"] = row.get("responsavelamostra", "")
        ws[f"H{idx}"] = row.get("responsavelcolheita", "")
        ws[f"I{idx}"] = ""  # Observações
        ws[f"K{idx}"] = row.get("procedure", "")
        ws[f"L{idx}"] = f"=A{idx}+30"  # Data requerido = Data receção + 30 dias

        # Campos obrigatórios
        for col in ["A", "B", "C", "D", "E", "F", "G"]:
            cell = ws[f"{col}{idx}"]
            if not cell.value or str(cell.value).strip() == "":
                cell.fill = red_fill

        # Destaque amarelo (revisão)
        if row.get("WasCorrected") or row.get("ValidationStatus") in ("review", "unknown", "no_list"):
            ws[f"D{idx}"].fill = yellow_fill

    # ────────────────────────────────────────────────
    # 📊 Validação E1:F1 — Nº Amostras declaradas / processadas
    # ────────────────────────────────────────────────
    processed = len(ocr_rows)
    expected = expected_count if expected_count is not None else "?"
    ws.merge_cells("E1:F1")
    cell = ws["E1"]
    cell.value = f"Nº Amostras: {expected} / {processed}"
    cell.font = bold_center
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.fill = green_fill if expected == processed else red_fill

    # ────────────────────────────────────────────────
    # 🗂️ Origem do PDF
    # ────────────────────────────────────────────────
    ws.merge_cells("G1:J1")
    pdf_orig_name = os.path.basename(source_pdf) if source_pdf else "(desconhecida)"
    ws["G1"].value = f"Origem: {pdf_orig_name}"
    ws["G1"].font = Font(italic=True, color="555555")
    ws["G1"].alignment = Alignment(horizontal="left", vertical="center")
    ws["G1"].fill = gray_fill

    # ────────────────────────────────────────────────
    # 🕒 Data/hora de processamento
    # ────────────────────────────────────────────────
    ws.merge_cells("K1:L1")
    timestamp = datetime.now().strftime("%d/%m/%Y %H:%M")
    ws["K1"].value = f"Processado em: {timestamp}"
    ws["K1"].font = Font(italic=True, color="555555")
    ws["K1"].alignment = Alignment(horizontal="right", vertical="center")
    ws["K1"].fill = gray_fill

    # 💾 Guardar ficheiro
    base_name = os.path.splitext(os.path.basename(out_name))[0]
    out_path = os.path.join(OUTPUT_DIR, f"{base_name}.xlsx")
    wb.save(out_path)
    print(f"🟢 Gravado com sucesso: {out_path}")

    return out_path

# ───────────────────────────────────────────────
# OCR + Parsing (devolve listas por requisição)
# ───────────────────────────────────────────────
# ──────────────────────────────────────────────────────────────────────────────
# 6. UTILITÁRIOS GERAIS
# ──────────────────────────────────────────────────────────────────────────────
def clean_value(s: str) -> str:
    """Limpa e normaliza um valor OCR."""
    if s is None:
        return ""
    if isinstance(s, (int, float)):
        return str(s)
    s = re.sub(r"[\u200b\t\r\f\v]+", " ", s)
    s = (s.strip()
           .replace("N/A", "")
           .replace("%", "")
           .replace("\n", " ")
           .replace("  ", " "))
    return s.strip()

def to_datetime(value: str):
    try:
        return datetime.strptime(str(value).strip(), "%d/%m/%Y")
    except Exception:
        return None

def pdf_to_images(pdf_path):
    """Converte PDF em imagens."""
    return convert_from_path(pdf_path, dpi=150)
# ───────────────────────────────────────────────
# 11. FUNÇÕES ASSÍNCRONAS DE OCR (Azure)
# ───────────────────────────────────────────────

async def azure_ocr_page(session, img_bytes, page_idx, pdf_name, cache_dir):
    """Envia imagem para OCR Azure (com cache e ignorando páginas vazias)."""
    cache_file = os.path.join(cache_dir, f"{os.path.basename(pdf_name)}_p{page_idx}.json")

    # Cache: evita chamadas repetidas
    if os.path.exists(cache_file):
        with open(cache_file, "r", encoding="utf-8") as f:
            result_json = json.load(f)
        print(f"📦 OCR cache usado (página {page_idx})")
    else:
        url = f"{AZURE_ENDPOINT}formrecognizer/documentModels/{MODEL_ID}:analyze?api-version=2023-07-31"
        headers = {"Ocp-Apim-Subscription-Key": AZURE_API_KEY, "Content-Type": "application/octet-stream"}

        async with session.post(url, data=img_bytes, headers=headers) as resp:
            if resp.status != 202:
                txt = await resp.text()
                print(f"❌ Erro Azure ({page_idx}): {txt}")
                return None
            result_url = resp.headers.get("Operation-Location")

        # Polling até o OCR estar concluído
        for _ in range(20):
            await asyncio.sleep(2)
            async with session.get(result_url, headers={"Ocp-Apim-Subscription-Key": AZURE_API_KEY}) as r:
                j = await r.json()
                if j.get("status") == "succeeded":
                    result_json = j
                    break
        else:
            print(f"⚠️ Timeout OCR página {page_idx}")
            return None

        with open(cache_file, "w", encoding="utf-8") as f:
            json.dump(result_json, f, ensure_ascii=False)

    # Ignorar páginas em branco
    text = extract_all_text(result_json)
    if len(text.strip()) < 100:
        return None

    tables = result_json.get("analyzeResult", {}).get("tables", [])
    return (page_idx, text, tables)

# ───────────────────────────────────────────────
#  OCR + PARSING COMPLETO PARA UM PDF
# ───────────────────────────────────────────────

async def process_pdf_async(pdf_path, session):
    print(f"\n📄 A processar async: {os.path.basename(pdf_path)}")
    t0 = asyncio.get_event_loop().time()

    cache_dir = os.path.join(OUTPUT_DIR, "_ocr_cache")
    os.makedirs(cache_dir, exist_ok=True)

    # Converter PDF em imagens (thread pool)
    with ThreadPoolExecutor() as pool:
        images = await asyncio.get_event_loop().run_in_executor(pool, lambda: pdf_to_images(pdf_path))

    # OCR paralelo das páginas
    tasks = []
    for i, img in enumerate(images, start=1):
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        img_bytes = buf.getvalue()
        tasks.append(azure_ocr_page(session, img_bytes, i, pdf_path, cache_dir))

    results = await asyncio.gather(*tasks)
    results = [r for r in results if r]

    if not results:
        print("⚠️ Nenhuma página útil após OCR.")
        return

    results.sort(key=lambda x: x[0])
    full_text = "\n".join([f"\n\n--- PÁGINA {i} ---\n{text}" for i, text, _ in results])
    all_tables = [t for _, _, tbls in results for t in tbls]

    t_ocr = asyncio.get_event_loop().time() - t0
    print(f"⏱️ OCR paralelo total: {timedelta(seconds=round(t_ocr))}")

    # Guardar texto global
    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
    txt_path = os.path.join(OUTPUT_DIR, base_name + "_ocr_debug.txt")
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write(full_text)

    # Parsing normal (usa função já existente)
    t_parse_start = asyncio.get_event_loop().time()
    combined_json = {"analyzeResult": {"tables": all_tables, "pages": []}}
    rows = parse_xylella_from_result(combined_json, pdf_path, txt_path)
    t_parse = asyncio.get_event_loop().time() - t_parse_start

    print(f"⏱️ Parsing + geração Excel: {timedelta(seconds=round(t_parse))}")
    print(f"🏁 Total: {timedelta(seconds=round(asyncio.get_event_loop().time() - t0))}")

    return rows


# ───────────────────────────────────────────────
#  PROCESSAMENTO ASSÍNCRONO DE TODOS OS PDFs
# ───────────────────────────────────────────────

async def process_folder_async(input_dir):
    """
    Processa todos os PDFs de forma assíncrona e gera um resumo final
    com tempos médios, totais e nº de amostras extraídas.
    """
    pdfs = [os.path.join(input_dir, f) for f in os.listdir(input_dir) if f.lower().endswith(".pdf")]
    if not pdfs:
        print("ℹ️ Não há PDFs na pasta de entrada.")
        return

    start_time = asyncio.get_event_loop().time()
    summary = []

    async with aiohttp.ClientSession() as session:
        tasks = [process_pdf_async(pdf, session) for pdf in pdfs]
        results = await asyncio.gather(*tasks)

    # Montar resumo de desempenho
    total_time = asyncio.get_event_loop().time() - start_time
    total_pdfs = len(pdfs)
    total_rows = 0
    total_time_ocr = 0
    total_time_parse = 0

    # Cada resultado é o "rows" retornado por process_pdf_async
    for pdf, res in zip(pdfs, results):
        if not res:
            continue
        total_rows += len(res)

    avg_time_per_pdf = total_time / total_pdfs if total_pdfs else 0

    print("\n📊 Resumo Final")
    print("──────────────────────────────")
    print(f"📄 PDFs processados: {total_pdfs}")
    print(f"🧾 Total de amostras extraídas: {total_rows}")
    print(f"⏱️ Tempo total: {timedelta(seconds=round(total_time))}")
    print(f"⚙️ Tempo médio por PDF: {timedelta(seconds=round(avg_time_per_pdf))}")
    print(f"📂 Saída: {OUTPUT_DIR}")
    print("──────────────────────────────\n")

# ───────────────────────────────────────────────
# API pública — processamento síncrono
# ───────────────────────────────────────────────
def process_pdf_sync(pdf_path: str):
    """
    Executa o OCR Azure e o parser Colab de forma síncrona.
    Devolve listas de amostras por requisição.
    """
    print(f"\n🧪 Início de processamento: {os.path.basename(pdf_path)}")

    # 1️⃣ OCR Azure
    result_json = azure_analyze_pdf(pdf_path)

    # 2️⃣ Gera texto OCR detalhado (para permitir deteção de cabeçalhos)
    base = os.path.splitext(os.path.basename(pdf_path))[0]
    txt_path = OUTPUT_DIR / f"{base}_ocr_debug.txt"

    full_text = extract_all_text(result_json)
    txt_path.write_text(full_text, encoding="utf-8")

    if len(full_text) < 2000:
        print("⚠️ OCR curto — pode não conter todos os cabeçalhos. Verifica se o PDF tem imagens digitalizadas.")

    # 3️⃣ Parser completo (com deteção de múltiplas requisições)
    rows_per_req = parse_all_requisitions(result_json, pdf_path, str(txt_path))

    # 4️⃣ Estatísticas finais
    total_amostras = sum(len(r) for r in rows_per_req)
    print(f"✅ {os.path.basename(pdf_path)}: {len(rows_per_req)} requisições, {total_amostras} amostras extraídas.")

    return rows_per_req

pass







