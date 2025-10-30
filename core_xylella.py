# -*- coding: utf-8 -*-
"""
core_xylella.py — Motor principal (Cloud/Streamlit)

Responsável por:
 - OCR via Azure Form Recognizer (PDF direto) com fallback Tesseract
 - Deteção de múltiplas requisições (DGAV)
 - Parsing de amostras a partir de TABELAS Azure (parser validado do Colab)
 - Export para TEMPLATE_PXf_SGSLABIP1056.xlsx
 - Logging: process_log.csv e resumo diário

API esperada pela UI:
    • process_pdf_sync(pdf_path) -> list[list[dict]]
    • write_to_template(rows_per_req, out_base_path, expected_count=None, source_pdf=None)
"""

from __future__ import annotations

import os, io, re, json, time
from pathlib import Path
from datetime import datetime, timedelta
from typing import List, Dict, Any, Tuple

import requests
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Alignment

# ───────────────────────────────────────────────────────────────
# Caminhos e configuração
# ───────────────────────────────────────────────────────────────
BASE_DIR = Path(__file__).parent
OUTPUT_DIR = Path(os.environ.get("OUTPUT_DIR", BASE_DIR / "Output"))
OUTPUT_DIR.mkdir(exist_ok=True)

TEMPLATE_PATH = Path(os.environ.get("TEMPLATE_PATH", BASE_DIR / "TEMPLATE_PXf_SGSLABIP1056.xlsx"))

AZURE_API_KEY = os.environ.get("AZURE_API_KEY", "")
AZURE_ENDPOINT = os.environ.get("AZURE_ENDPOINT", "")  # ex: "https://<nome>.cognitiveservices.azure.com/"
MODEL_ID = os.environ.get("AZURE_MODEL_ID", "prebuilt-document")

# ───────────────────────────────────────────────────────────────
# Estilos (Excel)
# ───────────────────────────────────────────────────────────────
YELLOW = PatternFill(start_color="FFFACD", end_color="FFFACD", fill_type="solid")
GREEN  = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
RED    = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
GRAY   = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
BOLD   = Font(bold=True, color="000000")
ITALIC = Font(italic=True, color="555555")

# ───────────────────────────────────────────────────────────────
# Utilitários
# ───────────────────────────────────────────────────────────────
def _now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def _clean(s: str) -> str:
    if s is None: return ""
    return re.sub(r"[\u200b\t\r\f\v]+", " ", str(s)).strip()

def _is_valid_date(value: Any) -> bool:
    if isinstance(value, datetime):
        return True
    try:
        datetime.strptime(str(value).strip(), "%d/%m/%Y")
        return True
    except Exception:
        return False

def _to_datetime(value: Any) -> datetime | None:
    try:
        return datetime.strptime(str(value).strip(), "%d/%m/%Y")
    except Exception:
        return None

# ───────────────────────────────────────────────────────────────
# OCR — Azure (PDF direto) + fallback Tesseract
# ───────────────────────────────────────────────────────────────
def azure_analyze_pdf(pdf_path: str) -> Dict[str, Any]:
    """Envia o PDF para o Azure Form Recognizer e retorna o JSON final."""
    if not AZURE_API_KEY or not AZURE_ENDPOINT:
        raise RuntimeError("Azure não configurado (AZURE_API_KEY/AZURE_ENDPOINT em falta).")

    url = f"{AZURE_ENDPOINT.rstrip('/')}/formrecognizer/documentModels/{MODEL_ID}:analyze?api-version=2023-07-31"
    headers = {"Ocp-Apim-Subscription-Key": AZURE_API_KEY, "Content-Type": "application/pdf"}

    with open(pdf_path, "rb") as f:
        resp = requests.post(url, data=f.read(), headers=headers, timeout=60)
    if resp.status_code != 202:
        raise RuntimeError(f"Azure analyze falhou: {resp.status_code} {resp.text}")

    op = resp.headers.get("Operation-Location")
    if not op:
        raise RuntimeError("Azure não devolveu Operation-Location.")

    # Polling
    for _ in range(40):
        time.sleep(1.5)
        r = requests.get(op, headers={"Ocp-Apim-Subscription-Key": AZURE_API_KEY}, timeout=30)
        j = r.json()
        if j.get("status") == "succeeded":
            return j
        if j.get("status") == "failed":
            raise RuntimeError(f"OCR Azure falhou: {j}")
    raise RuntimeError("Timeout a aguardar OCR Azure.")

def tesseract_fallback_text(pdf_path: str) -> str:
    """Fallback simples: extrai texto via Tesseract página a página (sem tabelas)."""
    try:
        import fitz  # PyMuPDF
        import pytesseract
        from PIL import Image
    except Exception as e:
        raise RuntimeError(f"Fallback Tesseract indisponível: {e}")

    doc = fitz.open(pdf_path)
    chunks = []
    for i in range(len(doc)):
        page = doc.load_page(i)
        pix = page.get_pixmap(dpi=200)
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        txt = pytesseract.image_to_string(img) or ""
        chunks.append(f"\n\n--- PÁGINA {i+1} ---\n{txt}")
    return "\n".join(chunks)

def extract_all_text(result_json: Dict[str, Any]) -> str:
    lines = []
    for pg in result_json.get("analyzeResult", {}).get("pages", []):
        for ln in pg.get("lines", []):
            lines.append(ln.get("content", "") or ln.get("text", ""))
    return "\n".join([_clean(x) for x in lines if _clean(x)])

# ───────────────────────────────────────────────────────────────
# Parser — constantes (como no Colab)
# ───────────────────────────────────────────────────────────────
NATUREZA_KEYWORDS = [
    "ramos", "folhas", "ramosefolhas", "ramosc/folhas",
    "material", "materialherbalho", "materialherbário", "materialherbalo",
    "natureza", "insetos", "sementes", "solo"
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

# ───────────────────────────────────────────────────────────────
# Deteção / Split de Requisições (como no Colab)
# ───────────────────────────────────────────────────────────────
def detect_requisicoes(full_text: str) -> Tuple[int, list[int]]:
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

def split_if_multiple_requisicoes(full_text: str) -> List[str]:
    text = re.sub(r"[ \t]+", " ", full_text)
    text = re.sub(r"\n{2,}", "\n", text)
    pattern = re.compile(
        r"(?:PROGRAMA\s+NACIONAL\s+DE\s+PROSPE[ÇC][AÃ]O\s+DE\s+PRAGAS\s+DE\s+QUARENTENA)",
        re.IGNORECASE,
    )
    marks = [m.start() for m in pattern.finditer(text)]
    if not marks:
        print("🔍 Nenhum cabeçalho encontrado — tratado como 1 requisição.")
        return [text]
    if len(marks) == 1:
        print("🔍 Apenas 1 cabeçalho — 1 requisição detectada.")
        return [text]
    marks.append(len(text))
    blocos = []
    for i in range(len(marks) - 1):
        start = max(0, marks[i] - 200)
        end = min(len(text), marks[i + 1] + 200)
        bloco = text[start:end].strip()
        if len(bloco) > 400:
            blocos.append(bloco)
        else:
            print(f"⚠️ Bloco {i+1} demasiado pequeno ({len(bloco)} chars) — possivelmente OCR truncado.")
    print(f"🔍 Detetadas {len(blocos)} requisições distintas (por cabeçalho).")
    return blocos

# ───────────────────────────────────────────────────────────────
# Contexto Global (como no Colab)
# ───────────────────────────────────────────────────────────────
def extract_context_from_text(full_text: str) -> Dict[str, Any]:
    ctx: Dict[str, Any] = {}
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
        dgav = re.sub(r"[:;,.\-–—]+$", "", m_d.group(0)).strip() if m_d else "DGAV"

    ctx["dgav"] = dgav
    ctx["responsavel_colheita"] = None

    # Datas de colheita por marcas (*), (**)
    colheita_map: Dict[str, str] = {}
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

    # Nº amostras declaradas
    flat = re.sub(r"\s+", " ", full_text)
    m_decl = re.search(r"N[º°]?\s*de\s*amostras(?:\s+neste\s+envio)?\s*[:\-]?\s*(\d{1,4})", flat, re.I)
    ctx["declared_samples"] = int(m_decl.group(1)) if m_decl else None

    return ctx

# ───────────────────────────────────────────────────────────────
# Parser de Tabelas (como no Colab)
# ───────────────────────────────────────────────────────────────
def parse_xylella_tables(result_json: Dict[str, Any], context: Dict[str, Any], req_id=None) -> List[Dict[str, Any]]:
    out: List[Dict[str, Any]] = []
    tables = result_json.get("analyzeResult", {}).get("tables", [])
    if not tables:
        print("⚠️ Nenhuma tabela encontrada.")
        return out

    def clean_value(x: Any) -> str:
        if x is None: return ""
        if isinstance(x, (int, float)): return str(x)
        s = re.sub(r"[\u200b\t\r\f\v]+", " ", str(x))
        return re.sub(r"\s{2,}", " ", s.replace("\n", " ").strip())

    for t in tables:
        nc = max((c.get("columnIndex", 0) for c in t.get("cells", [])), default=-1) + 1
        nr = max((c.get("rowIndex", 0) for c in t.get("cells", [])), default=-1) + 1
        grid = [[""] * nc for _ in range(max(nr, 0))]
        for c in t.get("cells", []):
            r, cidx = c.get("rowIndex", 0), c.get("columnIndex", 0)
            if r < 0 or cidx < 0: 
                continue
            if r >= len(grid):
                grid.extend([[""] * nc for _ in range(r - len(grid) + 1)])
            grid[r][cidx] = clean_value(c.get("content", ""))

        for row in grid:
            if not row or not any(row): 
                continue
            ref_raw = (row[0] if len(row) > 0 else "").strip()
            ref = _clean_ref(ref_raw)
            if not ref or re.match(r"^\D+$", ref):
                continue

            hospedeiro = row[2] if len(row) > 2 else ""
            obs        = row[3] if len(row) > 3 else ""

            if _looks_like_natureza(hospedeiro):
                hospedeiro = ""

            tipo = ""
            joined = " ".join([x for x in row if isinstance(x, str)])
            m_tipo = TIPO_RE.search(joined)
            if m_tipo:
                tipo = m_tipo.group(1).capitalize()
                obs = re.sub(TIPO_RE, "", obs).strip()

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

# ───────────────────────────────────────────────────────────────
# Parsing completo (devolve listas por requisição)
# ───────────────────────────────────────────────────────────────
def parse_all_requisitions(result_json: Dict[str, Any], pdf_name: str, txt_path: str) -> List[List[Dict[str, Any]]]:
    # Texto global (para split e contexto)
    if txt_path and os.path.exists(txt_path):
        with open(txt_path, "r", encoding="utf-8") as f:
            full_text = f.read()
        print(f"📝 Contexto extraído de {os.path.basename(txt_path)}")
    else:
        full_text = extract_all_text(result_json)

    count, _ = detect_requisicoes(full_text)
    all_tables = result_json.get("analyzeResult", {}).get("tables", []) or []

    if count <= 1:
        print("📄 Documento contém apenas uma requisição.")
        context = extract_context_from_text(full_text)
        amostras = parse_xylella_tables(result_json, context, req_id=1)
        return [amostras] if amostras else []

    blocos = split_if_multiple_requisicoes(full_text)
    print(f"📄 Documento dividido em {len(blocos)} requisições.")
    out: List[List[Dict[str, Any]]] = []

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

            tables_filtradas = [
                t for t in all_tables
                if any(ref in " ".join(c.get("content", "") for c in t.get("cells", []))
                       for ref in refs_bloco)
            ] or all_tables

            result_local = {"analyzeResult": {"tables": tables_filtradas}}
            amostras = parse_xylella_tables(result_local, context, req_id=i)
            if amostras:
                out.append(amostras)
            else:
                print(f"⚠️ Requisição {i} sem amostras válidas.")
        except Exception as e:
            print(f"❌ Erro na requisição {i}: {e}")

    return out

# ───────────────────────────────────────────────────────────────
# Logging (CSV + resumo diário)
# ───────────────────────────────────────────────────────────────
def append_process_log(pdf_name: str, req_id: int, processed: int, expected: int | None,
                       out_path: str | None, status="OK", error_msg: str | None = None) -> None:
    import csv
    log_path = OUTPUT_DIR / "process_log.csv"
    today_str = datetime.now().strftime("%Y-%m-%d")
    summary_path = OUTPUT_DIR / f"process_summary_{today_str}.txt"

    file_exists = log_path.exists()
    with open(log_path, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f, delimiter=";")
        if not file_exists:
            writer.writerow(["DataHora", "PDF", "ReqID", "Processadas", "Requisitadas", "OutputExcel", "Status", "Mensagem"])
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        writer.writerow([timestamp, os.path.basename(pdf_name), req_id, processed, expected or "", out_path or "", status, error_msg or ""])

    try:
        with open(summary_path, "a", encoding="utf-8") as f:
            f.write(f"[{datetime.now().strftime('%H:%M:%S')}] {os.path.basename(pdf_name)} (req {req_id}) → {status} "
                    f"({processed}/{expected or '?'} amostras) {os.path.basename(out_path or '')}\n")
    except Exception as e:
        print(f"⚠️ Falha ao atualizar resumo diário: {e}")

# ───────────────────────────────────────────────────────────────
# API — Escrever no TEMPLATE (um Excel por requisição)
# ───────────────────────────────────────────────────────────────
def write_to_template(rows_per_req, out_base_path, expected_count=None, source_pdf=None):
    """
    Escreve listas de amostras (uma por requisição) no TEMPLATE SGS.
    Gera vários ficheiros: <out_base>_req1.xlsx, req2.xlsx, ...
    Se vier 0 amostras/0 requisições, cria 1 ficheiro vazio com metadados.
    """
    from openpyxl import load_workbook

    template_path = Path(os.environ.get("TEMPLATE_PATH", TEMPLATE_PATH))
    if not template_path.exists():
        raise FileNotFoundError(f"TEMPLATE não encontrado: {template_path}")

    output_dir = Path(os.environ.get("OUTPUT_DIR", OUTPUT_DIR))
    output_dir.mkdir(exist_ok=True)

    sheet_name = "Avaliação pré registo"
    start_row = 6
    base_name = Path(out_base_path).stem
    out_files = []

    # garante que temos pelo menos 1 “requisição” para gravar metadados
    effective_reqs = rows_per_req if rows_per_req else [[]]

    for idx, req_rows in enumerate(effective_reqs, start=1):
        out_path = output_dir / f"{base_name}_req{idx}.xlsx"
        wb = load_workbook(template_path)
        if sheet_name not in wb.sheetnames:
            wb.close()
            raise KeyError(f"Folha '{sheet_name}' não encontrada no template.")
        ws = wb[sheet_name]

        # limpar zona de dados (sem tocar em cabeçalhos/fórmulas)
        max_lines = max(len(req_rows), 1) + 5
        for r in range(start_row, start_row + max_lines):
            for c in range(1, 13):
                ws.cell(row=r, column=c).value = None

        # escrever linhas
        for ridx, row in enumerate(req_rows, start=start_row):
            rececao_val = row.get("datarececao", "")
            colheita_val = row.get("datacolheita", "")

            def _is_valid_date(v):
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

            A = ws[f"A{ridx}"]; B = ws[f"B{ridx}"]
            A.value = _to_dt(rececao_val) if _is_valid_date(rececao_val) else rececao_val
            B.value = _to_dt(colheita_val) if _is_valid_date(colheita_val) else colheita_val

            ws[f"C{ridx}"] = row.get("referencia", "")
            ws[f"D{ridx}"] = row.get("hospedeiro", "")
            ws[f"E{ridx}"] = row.get("tipo", "")
            ws[f"F{ridx}"] = row.get("zona", "")
            ws[f"G{ridx}"] = row.get("responsavelamostra", "")
            ws[f"H{ridx}"] = row.get("responsavelcolheita", "")
            ws[f"I{ridx}"] = ""  # Observações
            ws[f"K{ridx}"] = row.get("procedure", "XYLELLA")
            ws[f"L{ridx}"] = f"=A{ridx}+30"  # Data requerido

        # contagem processada (mesmo 0)
        processed = len(req_rows)

        # E1:F1 — validação nº amostras
        ws.merge_cells("E1:F1")
        ws["E1"].value = f"Nº Amostras: {(expected_count if expected_count is not None else '?')} / {processed}"
        ws["E1"].font = Font(bold=True, color="000000")
        ws["E1"].alignment = Alignment(horizontal="center", vertical="center")
        ws["E1"].fill = (PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                         if (expected_count is not None and expected_count != processed)
                         else PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"))

        # G1:J1 — origem
        ws.merge_cells("G1:J1")
        ws["G1"].value = f"Origem: {os.path.basename(source_pdf) if source_pdf else base_name}"
        ws["G1"].font = Font(italic=True, color="555555")
        ws["G1"].alignment = Alignment(horizontal="left", vertical="center")
        ws["G1"].fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")

        # K1:L1 — timestamp
        ws.merge_cells("K1:L1")
        ws["K1"].value = f"Processado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        ws["K1"].font = Font(italic=True, color="555555")
        ws["K1"].alignment = Alignment(horizontal="right", vertical="center")
        ws["K1"].fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")

        wb.save(out_path)
        wb.close()
        print(f"🟢 Gravado (E1/F1, G1:J1, K1/L1): {out_path}")
        out_files.append(str(out_path))

    return out_files


# ───────────────────────────────────────────────────────────────
# API — Processar PDF (OCR + parsing) → listas por requisição
# ───────────────────────────────────────────────────────────────
def process_pdf_sync(pdf_path: str) -> List[List[Dict[str, Any]]]:
    pdf_path = str(pdf_path)
    base = Path(pdf_path).stem

    # OCR Azure
    try:
        print(f"📄 A processar (Azure): {os.path.basename(pdf_path)}")
        result_json = azure_analyze_pdf(pdf_path)
        full_text = extract_all_text(result_json)
    except Exception as e:
        print(f"⚠️ Azure falhou ({e}) — fallback Tesseract (texto simples).")
        full_text = tesseract_fallback_text(pdf_path)
        # quando só há texto, não temos tabelas — devolvemos parsing mínimo
        # (mantemos compatibilidade retornando estruturas por requisição)
        blocos = split_if_multiple_requisicoes(full_text)
        out_min: List[List[Dict[str, Any]]] = []
        for i, bloco in enumerate(blocos, start=1):
            ctx = extract_context_from_text(bloco)
            # regex simples: tenta apanhar referências tipo 123/2025/LVT/1 ou 7-8 dígitos
            refs = re.findall(r"\b(\d{2,4}/\d{2,4}/[A-Z0-9\-]+|\d{7,8})\b", bloco)
            req_rows = []
            for ref in refs:
                req_rows.append({
                    "requisicao_id": i,
                    "datarececao": ctx["data_envio"],
                    "datacolheita": ctx.get("default_colheita",""),
                    "referencia": _clean_ref(ref),
                    "hospedeiro": "",
                    "tipo": "",
                    "zona": ctx["zona"],
                    "responsavelamostra": ctx["dgav"],
                    "responsavelcolheita": ctx["responsavel_colheita"],
                    "observacoes": "",
                    "procedure": "XYLELLA",
                    "datarequerido": ctx["data_envio"],
                    "Score": ""
                })
            if req_rows:
                out_min.append(req_rows)
        # guardar debug
        debug_txt = OUTPUT_DIR / f"{base}_ocr_debug.txt"
        with open(debug_txt, "w", encoding="utf-8") as f:
            f.write(full_text)
        print(f"📝 Texto OCR bruto guardado em: {debug_txt}")
        return out_min

    # Guardar debug OCR bruto (Azure)
    debug_txt = OUTPUT_DIR / f"{base}_ocr_debug.txt"
    with open(debug_txt, "w", encoding="utf-8") as f:
        f.write(full_text)
    print(f"📝 Texto OCR bruto guardado em: {debug_txt}")

    # Parsing completo (como no Colab) — listas por requisição
    rows_per_req = parse_all_requisitions(result_json, pdf_path, str(debug_txt))
    total_amostras = sum(len(req) for req in rows_per_req)
    print(f"🏁 OCR+Parsing concluído: {len(rows_per_req)} requisições, {total_amostras} amostras.")
    return rows_per_req


# ───────────────────────────────────────────────────────────────
# Execução direta (debug local)
# ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="Xylella Processor (Azure OCR + Parser + Template SGS)")
    ap.add_argument("pdf", help="Caminho do PDF a processar")
    ap.add_argument("--expected", type=int, default=None, help="Nº de amostras esperado por requisição (E1/F1)")
    args = ap.parse_args()

    rows = process_pdf_sync(args.pdf)
    base = Path(args.pdf).stem
    files = write_to_template(rows, base, expected_count=args.expected, source_pdf=Path(args.pdf).name)

    print("\n📂 Saídas geradas:")
    for p in files:
        print(" -", p)

