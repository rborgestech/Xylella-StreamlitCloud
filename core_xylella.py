# -*- coding: utf-8 -*-
"""
core_xylella.py — Cloud/Streamlit (OCR Azure direto + Parser Colab + Writer por requisição)

API exposta e usada pela UI (xylella_processor.py):
    • process_pdf_sync(pdf_path) -> List[List[Dict]]]   # devolve lista de requisições; cada requisição = lista de amostras (dict)
    • write_to_template(rows, out_name, expected_count=None, source_pdf=None) -> str  # escreve 1 XLSX com base no template

Requer:
  - AZURE_API_KEY, AZURE_ENDPOINT (env)
  - TEMPLATE_PATH (env) ou ficheiro 'TEMPLATE_PXf_SGSLABIP1056.xlsx' ao lado do core
  - OUTPUT_DIR (env) — diretório onde guardar .xlsx e _ocr_debug.txt
"""

import os, re, time, tempfile, requests, csv
from datetime import datetime
from pathlib import Path
from typing import Dict, Any, List
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment

# ───────────────────────────────────────────────
# Diretórios e paths
# ───────────────────────────────────────────────
try:
    OUTPUT_DIR = Path(os.getenv("OUTPUT_DIR", tempfile.gettempdir()))
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
except Exception as e:
    OUTPUT_DIR = Path(tempfile.gettempdir())
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    print(f"[WARN] Não foi possível criar diretório de output definido: {e}. Usando {OUTPUT_DIR}")

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = Path(os.environ.get("TEMPLATE_PATH", BASE_DIR / "TEMPLATE_PXf_SGSLABIP1056.xlsx"))
if not TEMPLATE_PATH.exists():
    print(f"ℹ️ Aviso: TEMPLATE não encontrado em {TEMPLATE_PATH}. Será verificado no momento da escrita.")

# ───────────────────────────────────────────────
# Azure OCR — credenciais
# ───────────────────────────────────────────────
AZURE_API_KEY = os.environ.get("AZURE_API_KEY", "")
AZURE_ENDPOINT = os.environ.get("AZURE_ENDPOINT", "")
MODEL_ID = os.environ.get("AZURE_MODEL_ID", "prebuilt-document")

# ───────────────────────────────────────────────
# OCR Azure
# ───────────────────────────────────────────────
def azure_analyze_pdf(pdf_path: str) -> Dict[str, Any]:
    if not AZURE_API_KEY or not AZURE_ENDPOINT:
        raise RuntimeError("Azure não configurado (AZURE_API_KEY/AZURE_ENDPOINT).")

    url = f"{AZURE_ENDPOINT.rstrip('/')}/formrecognizer/documentModels/{MODEL_ID}:analyze?api-version=2023-07-31"
    headers = {"Ocp-Apim-Subscription-Key": AZURE_API_KEY, "Content-Type": "application/pdf"}

    with open(pdf_path, "rb") as f:
        resp = requests.post(url, data=f.read(), headers=headers, timeout=120)
    if resp.status_code != 202:
        raise RuntimeError(f"Azure analyze falhou: {resp.status_code} {resp.text}")

    op = resp.headers.get("Operation-Location")
    if not op:
        raise RuntimeError("Azure não devolveu Operation-Location.")

    start = time.time()
    while True:
        r = requests.get(op, headers={"Ocp-Apim-Subscription-Key": AZURE_API_KEY}, timeout=60)
        j = r.json()
        st = j.get("status")
        if st == "succeeded":
            return j
        if st == "failed":
            raise RuntimeError(f"OCR Azure falhou: {j}")
        if time.time() - start > 180:
            raise RuntimeError("Timeout a aguardar OCR Azure.")
        time.sleep(1.2)

# ───────────────────────────────────────────────
# Utilitários
# ───────────────────────────────────────────────
def extract_all_text(result_json: Dict[str, Any]) -> str:
    lines = []
    for pg in result_json.get("analyzeResult", {}).get("pages", []):
        for ln in pg.get("lines", []):
            txt = (ln.get("content") or ln.get("text") or "").strip()
            if txt:
                lines.append(txt)
    return "\n".join(lines)

def normalize_date_str(val: str) -> str:
    if not val:
        return ""
    txt = str(val).strip().replace("-", "/").replace(".", "/")
    txt = re.sub(r"[\u00A0\s]+", "", txt)
    m_std = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})$", txt)
    if m_std:
        d, m_, y = map(int, m_std.groups())
        if 1 <= d <= 31 and 1 <= m_ <= 12 and 1900 <= y <= 2100:
            return f"{d:02d}/{m_:02d}/{y:04d}"
    digits = re.sub(r"\D", "", txt)
    if len(digits) == 8:
        d, m_, y = int(digits[:2]), int(digits[2:4]), int(digits[4:])
        return f"{d:02d}/{m_:02d}/{y:04d}"
    return txt

def _to_datetime(value: str):
    try:
        norm = normalize_date_str(value)
        return datetime.strptime(norm, "%d/%m/%Y")
    except Exception:
        return None

# ───────────────────────────────────────────────
# Extração de contexto (zona, DGAV, datas, nº amostras)
# ───────────────────────────────────────────────
def extract_context_from_text(full_text: str):
    ctx = {}

    # Zona
    m_zona = re.search(r"Xylella\s+fastidiosa\s*\(([^)]+)\)", full_text, re.I)
    ctx["zona"] = m_zona.group(1).strip() if m_zona else "Zona Isenta"

    # DGAV
    m_dgav = re.search(r"\bDGAV(?:\s+[A-Za-zÀ-ÿ]+){1,4}", full_text)
    ctx["dgav"] = m_dgav.group(0).strip() if m_dgav else "DGAV"

    # Datas
    m_col = re.search(r"Data\s+de\s+colheita[:\-\s]*([0-9/\-\s]+)", full_text, re.I)
    ctx["default_colheita"] = normalize_date_str(m_col.group(1)) if m_col else ""
    m_envio = re.search(r"Data\s+(?:do|de)\s+envio.*?:?\s*([0-9/\-\s]+)", full_text, re.I)
    ctx["data_envio"] = normalize_date_str(m_envio.group(1)) if m_envio else ctx["default_colheita"]

    # ───────────────────────────────────────────────
    # Nº de amostras declaradas (robusto a OCR e placeholders)
    # ───────────────────────────────────────────────
    flat = re.sub(r"[\u00A0_\s]+", " ", full_text)  # normaliza espaços e underscores
    flat = flat.replace("–", "-").replace("—", "-")

    # aceita variações e ruído OCR
    patterns = [
        r"N[º°o]?\s*de\s*amostras(?:\s+neste\s+envio)?\s*[:\-]?\s*([0-9OoQIl]{1,4})\b",
        r"N\s*[º°o]?\s*amostras.*?([0-9OoQIl]{1,4})\b",
        r"amostras\s*(?:neste\s+envio)?\s*[:\-]?\s*([0-9OoQIl]{1,4})\b",
        r"n\s*o\s*de\s*amostras.*?([0-9OoQIl]{1,4})\b"
    ]
    found = None
    for pat in patterns:
        m_decl = re.search(pat, flat, re.I)
        if m_decl:
            found = m_decl.group(1)
            break

    if found:
        raw = found.strip()
        raw = (raw.replace("O", "0").replace("o", "0")
                    .replace("Q", "0").replace("q", "0")
                    .replace("I", "1").replace("l", "1"))
        try:
            ctx["declared_samples"] = int(raw)
        except ValueError:
            ctx["declared_samples"] = 0
    else:
        ctx["declared_samples"] = 0

    return ctx

# ───────────────────────────────────────────────
# Parser principal
# ───────────────────────────────────────────────
def parse_xylella_tables(result_json, context, req_id=None) -> List[Dict[str, Any]]:
    out: List[Dict[str, Any]] = []
    tables = result_json.get("analyzeResult", {}).get("tables", [])
    if not tables:
        print("⚠️ Nenhuma tabela encontrada.")
        return out

    for t in tables:
        nc = max(c.get("columnIndex", 0) for c in t.get("cells", [])) + 1
        nr = max(c.get("rowIndex", 0) for c in t.get("cells", [])) + 1
        grid = [[""] * nc for _ in range(nr)]
        for c in t.get("cells", []):
            grid[c["rowIndex"]][c["columnIndex"]] = (c.get("content") or "").strip()

        for row in grid:
            if not row or not any(row):
                continue
            ref = row[0].strip()
            if not ref or re.match(r"^\D+$", ref):
                continue

            out.append({
                "requisicao_id": req_id,
                "datarececao": context.get("data_envio"),
                "datacolheita": context.get("default_colheita"),
                "referencia": ref,
                "hospedeiro": row[2] if len(row) > 2 else "",
                "tipo": row[3] if len(row) > 3 else "",
                "zona": context["zona"],
                "responsavelamostra": context["dgav"],
                "responsavelcolheita": "",
                "procedure": "XYLELLA"
            })
    print(f"✅ {len(out)} amostras extraídas (req_id={req_id}).")
    return out

# ───────────────────────────────────────────────
# Escrita no template
# ───────────────────────────────────────────────
def write_to_template(ocr_rows, out_name, expected_count=None, source_pdf=None):
    if not ocr_rows:
        return None
    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(f"Template não encontrado: {TEMPLATE_PATH}")

    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.worksheets[0]
    start_row = 4

    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    gray_fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
    bold_center = Font(bold=True, color="000000")

    for row in range(start_row, 200):
        for col in range(1, 13):
            ws.cell(row=row, column=col).value = None

    for i, r in enumerate(ocr_rows, start=start_row):
        ws[f"A{i}"] = r.get("datarececao", "")
        ws[f"B{i}"] = r.get("datacolheita", "")
        ws[f"C{i}"] = r.get("referencia", "")
        ws[f"D{i}"] = r.get("hospedeiro", "")
        ws[f"E{i}"] = r.get("tipo", "")
        ws[f"F{i}"] = r.get("zona", "")
        ws[f"G{i}"] = r.get("responsavelamostra", "")
        ws[f"K{i}"] = "XYLELLA"

    ws.merge_cells("E1:F1")
    cell = ws["E1"]
    val_str = f"{len(ocr_rows)} / {expected_count or len(ocr_rows)}"
    cell.value = val_str
    cell.font = bold_center
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.fill = green_fill

    ws.merge_cells("G1:J1")
    ws["G1"].value = f"Origem: {os.path.basename(source_pdf) if source_pdf else ''}"
    ws["G1"].font = Font(italic=True, color="555555")
    ws["G1"].fill = gray_fill

    ws.merge_cells("K1:L1")
    ws["K1"].value = f"Processado em: {datetime.now():%d/%m/%Y %H:%M}"
    ws["K1"].font = Font(italic=True, color="555555")
    ws["K1"].fill = gray_fill

    out_path = OUTPUT_DIR / f"{os.path.splitext(out_name)[0]}.xlsx"
    wb.save(out_path)
    print(f"🟢 Gravado: {out_path}")
    return str(out_path)

# ───────────────────────────────────────────────
# Função principal
# ───────────────────────────────────────────────
def process_pdf_sync(pdf_path: str) -> List[str]:
    base = os.path.basename(pdf_path)
    print(f"\n🧪 Início de processamento: {base}")

    result_json = azure_analyze_pdf(pdf_path)
    text = extract_all_text(result_json)
    ctx = extract_context_from_text(text)
    amostras = parse_xylella_tables(result_json, ctx, req_id=1)
    expected = ctx.get("declared_samples", 0)
    out_name = f"{os.path.splitext(base)[0]}.xlsx"
    out_path = write_to_template(amostras, out_name, expected_count=expected, source_pdf=pdf_path)

    print(f"🏁 {base}: {len(amostras)} amostras processadas, esperado {expected}.")
    return [out_path]
