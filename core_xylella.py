# -*- coding: utf-8 -*-
"""
core_xylella.py — versão final (Cloud/Streamlit)

✔️ OCR Azure (PDF direto)
✔️ Parser Colab integrado
✔️ Um ficheiro Excel por requisição
✔️ Limpeza real do template (sem cache)
✔️ Total de amostras visível (E1/F1)
"""

from __future__ import annotations
import os, re, io, time, json, shutil, requests
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Any

from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment

# ───────────────────────────────────────────────
# Configuração de caminhos
# ───────────────────────────────────────────────
BASE_DIR = Path(__file__).parent
OUTPUT_DIR = Path(os.environ.get("OUTPUT_DIR", BASE_DIR / "Output"))
OUTPUT_DIR.mkdir(exist_ok=True)
TEMPLATE_PATH = Path(os.environ.get("TEMPLATE_PATH", BASE_DIR / "TEMPLATE_PXf_SGSLABIP1056.xlsx"))

AZURE_API_KEY = os.environ.get("AZURE_API_KEY", "")
AZURE_ENDPOINT = os.environ.get("AZURE_ENDPOINT", "")
MODEL_ID = os.environ.get("AZURE_MODEL_ID", "prebuilt-document")

# ───────────────────────────────────────────────
# Estilos Excel
# ───────────────────────────────────────────────
GREEN = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
RED = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
GRAY = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
BOLD = Font(bold=True, color="000000")
ITALIC = Font(italic=True, color="555555")

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
# OCR Azure
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

    for _ in range(50):
        time.sleep(1.2)
        r = requests.get(op, headers={"Ocp-Apim-Subscription-Key": AZURE_API_KEY}, timeout=30)
        j = r.json()
        if j.get("status") == "succeeded":
            return j
        if j.get("status") == "failed":
            raise RuntimeError(f"OCR Azure falhou: {j}")
    raise RuntimeError("Timeout a aguardar OCR Azure.")

# ───────────────────────────────────────────────
# Parser Colab (simplificado)
# ───────────────────────────────────────────────
TIPO_RE = re.compile(r"\b(Simples|Composta|Individual)\b", re.I)

def _clean_ref(raw: str) -> str:
    s = (raw or "").strip()
    s = re.sub(r"\s*/\s*", "/", s)
    s = re.sub(r"/{2,}", "/", s)
    s = re.sub(r"[A-Za-z]+", lambda m: m.group(0).upper(), s)
    s = re.sub(r"[^A-Z0-9/]+$", "", s)
    return s

def detect_requisicoes(full_text: str):
    pat = re.compile(r"PROGRAMA\s+NACIONAL\s+DE\s+PROSPE", re.I)
    marks = [m.start() for m in pat.finditer(full_text)]
    if not marks:
        return 1, []
    return len(marks), marks

def split_if_multiple_requisicoes(full_text: str):
    text = re.sub(r"[ \t]+", " ", full_text)
    text = re.sub(r"\n{2,}", "\n", text)
    pat = re.compile(r"PROGRAMA\s+NACIONAL\s+DE\s+PROSPE", re.I)
    marks = [m.start() for m in pat.finditer(text)]
    if not marks or len(marks) == 1:
        return [text]
    marks.append(len(text))
    blocos = []
    for i in range(len(marks) - 1):
        blocos.append(text[marks[i]:marks[i+1]].strip())
    print(f"📄 Documento dividido em {len(blocos)} requisições.")
    return blocos

def extract_context_from_text(full_text: str):
    ctx = {}
    m_zona = re.search(r"Xylella\s+fastidiosa\s*\(([^)]+)\)", full_text, re.I)
    ctx["zona"] = m_zona.group(1).strip() if m_zona else "Zona Isenta"
    m_envio = re.search(r"Data\s+(?:do|de)\s+envio.*?([\d/]{8,10})", full_text, re.I)
    ctx["data_envio"] = m_envio.group(1) if m_envio else datetime.now().strftime("%d/%m/%Y")
    ctx["dgav"] = "DGAV"
    ctx["responsavel_colheita"] = ""
    ctx["default_colheita"] = ctx["data_envio"]
    ctx["colheita_map"] = {}
    return ctx

def parse_xylella_tables(result_json: Dict[str, Any], context: Dict[str, Any], req_id=None):
    out = []
    tables = result_json.get("analyzeResult", {}).get("tables", []) or []
    if not tables:
        return out
    for t in tables:
        cells = t.get("cells", [])
        if not cells:
            continue
        for c in cells:
            val = str(c.get("content", "")).strip()
            if re.match(r"^\d{7,8}$", val) or re.match(r"^\d{1,4}/\d{4}/[A-Z]{2,}", val):
                ref = _clean_ref(val)
                out.append({
                    "requisicao_id": req_id,
                    "datarececao": context["data_envio"],
                    "datacolheita": context["default_colheita"],
                    "referencia": ref,
                    "hospedeiro": "",
                    "tipo": "",
                    "zona": context["zona"],
                    "responsavelamostra": context["dgav"],
                    "responsavelcolheita": context["responsavel_colheita"],
                    "observacoes": "",
                    "procedure": "XYLELLA",
                    "datarequerido": context["data_envio"],
                    "Score": ""
                })
    print(f"✅ {len(out)} amostras extraídas (req_id={req_id}).")
    return out

def parse_all_requisitions(result_json, pdf_name, txt_path):
    if txt_path and os.path.exists(txt_path):
        full_text = Path(txt_path).read_text(encoding="utf-8")
    else:
        full_text = extract_all_text(result_json)

    blocos = split_if_multiple_requisicoes(full_text)
    tables_all = result_json.get("analyzeResult", {}).get("tables", []) or []
    all_reqs = []

    for i, bloco in enumerate(blocos, start=1):
        context = extract_context_from_text(bloco)
        refs_bloco = re.findall(r"\b\d{7,8}\b|\b\d{2,4}/\d{2,4}/[A-Z0-9\-]+\b", bloco, re.I)
        tables_filtradas = []
        for t in tables_all:
            joined = " ".join(c.get("content", "") for c in t.get("cells", []))
            if any(ref in joined for ref in refs_bloco):
                tables_filtradas.append(t)
        if not tables_filtradas:
            print(f"⚠️ Sem correspondência de tabelas na requisição {i}. Ignorado.")
            continue
        local = {"analyzeResult": {"tables": tables_filtradas}}
        amostras = parse_xylella_tables(local, context, req_id=i)
        if amostras:
            all_reqs.append(amostras)
    return all_reqs

# ───────────────────────────────────────────────
# Escrita no TEMPLATE (1 ficheiro por requisição)
# ───────────────────────────────────────────────
def write_to_template(rows_per_req, out_base_path, expected_count=None, source_pdf=None):
    template_path = Path(os.environ.get("TEMPLATE_PATH", TEMPLATE_PATH))
    if not template_path.exists():
        raise FileNotFoundError(f"TEMPLATE não encontrado: {template_path}")
    out_dir = Path(os.environ.get("OUTPUT_DIR", OUTPUT_DIR))
    out_dir.mkdir(exist_ok=True)

    sheet = "Avaliação pré registo"
    start_row = 6
    base = Path(out_base_path).stem
    out_files = []
    effective = rows_per_req if rows_per_req else [[]]

    for idx, req in enumerate(effective, start=1):
        tmp_template = out_dir / f"__tmp_req{idx}.xlsx"
        shutil.copy(template_path, tmp_template)
        wb = load_workbook(tmp_template)
        ws = wb[sheet]

        # limpar zona de dados (sem tocar fórmulas)
        for r in range(start_row, ws.max_row + 1):
            for c in range(1, 13):
                ws.cell(row=r, column=c).value = None

        # escrever linhas
        for ridx, row in enumerate(req, start=start_row):
            A, B = ws[f"A{ridx}"], ws[f"B{ridx}"]
            rece, colh = row.get("datarececao", ""), row.get("datacolheita", "")
            A.value = _to_dt(rece) if _is_valid_date(rece) else rece
            B.value = _to_dt(colh) if _is_valid_date(colh) else colh
            ws[f"C{ridx}"] = row.get("referencia", "")
            ws[f"D{ridx}"] = row.get("hospedeiro", "")
            ws[f"E{ridx}"] = row.get("tipo", "")
            ws[f"F{ridx}"] = row.get("zona", "")
            ws[f"G{ridx}"] = row.get("responsavelamostra", "")
            ws[f"H{ridx}"] = row.get("responsavelcolheita", "")
            ws[f"I{ridx}"] = ""
            ws[f"K{ridx}"] = row.get("procedure", "XYLELLA")
            ws[f"L{ridx}"] = f"=A{ridx}+30"

        processed = len(req)
        ws.merge_cells("E1:F1")
        ws["E1"].value = f"Nº Amostras: {processed}"
        ws["E1"].font = BOLD
        ws["E1"].alignment = Alignment(horizontal="center", vertical="center")
        ws["E1"].fill = GREEN if processed > 0 else RED

        ws.merge_cells("G1:J1")
        ws["G1"].value = f"Origem: {os.path.basename(source_pdf) if source_pdf else base}"
        ws["G1"].font = ITALIC
        ws["G1"].alignment = Alignment(horizontal="left", vertical="center")
        ws["G1"].fill = GRAY

        ws.merge_cells("K1:L1")
        ws["K1"].value = f"Processado em: {datetime.now():%d/%m/%Y %H:%M}"
        ws["K1"].font = ITALIC
        ws["K1"].alignment = Alignment(horizontal="right", vertical="center")
        ws["K1"].fill = GRAY

        out_path = out_dir / f"{base}_req{idx}.xlsx"
        wb.save(out_path)
        wb.close()
        tmp_template.unlink(missing_ok=True)
        print(f"🟢 Gravado: {out_path}")
        out_files.append(str(out_path))

    return out_files

# ───────────────────────────────────────────────
# OCR + Parsing
# ───────────────────────────────────────────────
def process_pdf_sync(pdf_path: str):
    pdf_path = str(pdf_path)
    base = Path(pdf_path).stem
    result = azure_analyze_pdf(pdf_path)
    full_text = extract_all_text(result)
    debug_txt = OUTPUT_DIR / f"{base}_ocr_debug.txt"
    with open(debug_txt, "w", encoding="utf-8") as f:
        f.write(full_text)
    print(f"📝 Texto OCR bruto guardado em: {debug_txt}")
    rows_per_req = parse_all_requisitions(result, pdf_path, str(debug_txt))
    total = sum(len(x) for x in rows_per_req)
    print(f"🏁 Concluído: {len(rows_per_req)} requisições, {total} amostras.")
    return rows_per_req

# Execução direta
if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf")
    args = ap.parse_args()
    rows = process_pdf_sync(args.pdf)
    base = Path(args.pdf).stem
    write_to_template(rows, base, source_pdf=args.pdf)
