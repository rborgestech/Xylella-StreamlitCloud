# core_xylella.py
# -*- coding: utf-8 -*-
"""
Xylella Processor – OCR + Parser + Excel Writer
Autor: Rosa Borges
Data: 29/10/2025

Execução local (VS Code) + OneDrive (orquestrado fora deste módulo):
 - Faz OCR via Azure Form Recognizer
 - Extrai tabelas e contexto das requisições DGAV
 - Gera ficheiros Excel com base em TEMPLATE_PXf_SGSLABIP1056.xlsx
 - Mantém processamento assíncrono (rápido)
 - Sem dependências de Google Colab / Poppler (usa PyMuPDF)
"""

import os
import re
import io
import csv
import json
import time
import asyncio
import traceback
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor
from typing import Dict, List, Tuple, Optional

import aiohttp
import requests
from PIL import Image
import fitz  # PyMuPDF
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Alignment

# ───────────────────────────────────────────────────────────────
# 0) CONFIG LOCAL (pastas + ficheiros auxiliares)
# ───────────────────────────────────────────────────────────────
CWD = os.getcwd()
INPUT_DIR = os.path.join(CWD, "Input")
OUTPUT_DIR = os.path.join(CWD, "Output")
os.makedirs(INPUT_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Caminhos configuráveis por ENV; caso contrário, ficheiros na raiz do projeto
TEMPLATE_PATH = os.getenv("TEMPLATE_PATH", os.path.join(CWD, "TEMPLATE_PXf_SGSLABIP1056.xlsx"))
PLANT_LIST_PATH = os.getenv("PLANT_LIST_PATH", os.path.join(CWD, "PlantList.txt"))
PLANT_LIST_EXTRA_PATH = os.getenv("PLANT_LIST_EXTRA_PATH", os.path.join(CWD, "PlantList_EXTRA.txt"))

# ───────────────────────────────────────────────────────────────
# 1) CONFIG AZURE (via variáveis de ambiente)
# ───────────────────────────────────────────────────────────────
AZURE_API_KEY = os.getenv("AZURE_KEY")  # define no ambiente/VSCode
AZURE_ENDPOINT = os.getenv("AZURE_ENDPOINT", "https://ifapprod.cognitiveservices.azure.com/")
MODEL_ID = os.getenv("AZURE_MODEL_ID", "prebuilt-document")
if not AZURE_API_KEY:
    raise RuntimeError("AZURE_KEY não definida. Define a variável de ambiente AZURE_KEY com a tua chave Azure.")

# ───────────────────────────────────────────────────────────────
# 2) UTILITÁRIOS GERAIS
# ───────────────────────────────────────────────────────────────
def clean_value(s: str) -> str:
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

def _to_datetime(value: str):
    try:
        return datetime.strptime(str(value).strip(), "%d/%m/%Y")
    except Exception:
        return None

def pdf_to_images_pymupdf(pdf_path: str) -> List[Image.Image]:
    """Renderiza cada página do PDF como PIL.Image sem necessidade de Poppler."""
    imgs: List[Image.Image] = []
    with fitz.open(pdf_path) as doc:
        for page in doc:
            # 150-200 dpi equivalentes: usa matrix de escala
            zoom = 2.0  # ~144-150 dpi a partir dos 72dpi base
            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat, alpha=False)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            imgs.append(img)
    return imgs

def extract_all_text(result_json: dict) -> str:
    lines = []
    for page in result_json.get("analyzeResult", {}).get("pages", []):
        for line in page.get("lines", []):
            lines.append(line.get("content", ""))
    return "\n".join(lines)

# ───────────────────────────────────────────────────────────────
# 3) OCR com Azure Form Recognizer (assíncrono + cache)
# ───────────────────────────────────────────────────────────────
async def azure_ocr_page(session: aiohttp.ClientSession, img_bytes: bytes, page_idx: int, pdf_name: str, cache_dir: str):
    cache_file = os.path.join(cache_dir, f"{os.path.basename(pdf_name)}_p{page_idx}.json")

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

        # Polling do resultado
        for _ in range(30):  # até ~60s
            await asyncio.sleep(2)
            async with session.get(result_url, headers=headers) as r:
                j = await r.json()
                if j.get("status") == "succeeded":
                    result_json = j
                    break
        else:
            print(f"⚠️ Timeout OCR página {page_idx}")
            return None

        with open(cache_file, "w", encoding="utf-8") as f:
            json.dump(result_json, f, ensure_ascii=False)

    text = extract_all_text(result_json)
    if len(text.strip()) < 80:
        return None
    tables = result_json.get("analyzeResult", {}).get("tables", [])
    return (page_idx, text, tables)

# ───────────────────────────────────────────────────────────────
# 4) PARSING XYLELLA (regras tal como no teu código)
# ───────────────────────────────────────────────────────────────
NATUREZA_KEYWORDS = [
    "ramos", "folhas", "ramosefolhas", "ramosc/folhas",
    "material", "materialherbalho", "materialherbário", "materialherbalo",
    "natureza", "insetos", "sementes", "solo"
]
REF_NUM_RE   = re.compile(r"\b\d{7,8}\b")
TIPO_RE      = re.compile(r"\b(Composta|Simples)\b", re.I)

def _clean_ref(raw: str) -> str:
    s = raw.strip()
    s = re.sub(r"\s*/\s*", "/", s)
    s = re.sub(r"/{2,}", "/", s)
    s = re.sub(r"[A-Za-z]+", lambda m: m.group(0).upper(), s)
    s = s.replace("LUT", "LVT")
    s = re.sub(r"\s+", "", s)
    s = re.sub(r"[^A-Z0-9/]+$", "", s)
    return s

def _looks_like_natureza(txt: str) -> bool:
    t = re.sub(r"\s+", "", txt.lower())
    return any(k in t for k in NATUREZA_KEYWORDS)

def detect_requisicoes(full_text: str):
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

def extract_context_from_text(full_text: str):
    ctx: Dict[str, Optional[str]] = {}

    # Zona
    m_zona = re.search(r"Xylella\s+fastidiosa\s*\(([^)]+)\)", full_text, re.I)
    ctx["zona"] = m_zona.group(1).strip() if m_zona else "Zona Isenta"

    # DGAV/Responsável
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

    # Datas de colheita (map por asteriscos)
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

    # Nº amostras declaradas
    flat = re.sub(r"\s+", " ", full_text)
    m_decl = re.search(r"N[º°]?\s*de\s*amostras(?:\s+neste\s+envio)?\s*[:\-]?\s*(\d{1,4})", flat, re.I)
    ctx["declared_samples"] = int(m_decl.group(1)) if m_decl else None
    return ctx

def parse_xylella_tables(result_json: dict, context: dict, req_id: Optional[int] = None) -> List[Dict]:
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

def write_to_template(ocr_rows: List[Dict], out_name: str, expected_count: Optional[int] = None, source_pdf: Optional[str] = None):
    if not ocr_rows:
        print(f"⚠️ {out_name}: sem linhas para escrever.")
        return None
    if not os.path.exists(TEMPLATE_PATH):
        raise FileNotFoundError(f"Template não encontrado: {TEMPLATE_PATH}")

    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.worksheets[0]
    start_row = 4

    yellow_fill = PatternFill(start_color="FFFACD", end_color="FFFACD", fill_type="solid")
    green_fill  = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    red_fill    = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    gray_fill   = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
    bold_center = Font(bold=True, color="000000")

    # Limpar linhas anteriores
    for row in range(start_row, 201):
        for col in range(1, 13):
            cell = ws.cell(row=row, column=col)
            cell.value = None
            cell.fill = PatternFill(fill_type=None)
        ws[f"I{row}"].value = None

    def _is_valid_date(value: str) -> bool:
        try:
            if isinstance(value, datetime):
                return True
            datetime.strptime(str(value).strip(), "%d/%m/%Y")
            return True
        except Exception:
            return False

    def _to_dt(value: str):
        try:
            return datetime.strptime(str(value).strip(), "%d/%m/%Y")
        except Exception:
            return None

    # Escrever linhas
    for idx, row in enumerate(ocr_rows, start=start_row):
        rececao_val   = row.get("datarececao", "")
        colheita_val  = row.get("datacolheita", "")

        cell_A = ws[f"A{idx}"]
        cell_B = ws[f"B{idx}"]
        if _is_valid_date(rececao_val):
            cell_A.value = _to_dt(rececao_val)
        else:
            cell_A.value = rececao_val
            cell_A.fill = red_fill

        if _is_valid_date(colheita_val):
            cell_B.value = _to_dt(colheita_val)
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
        ws[f"L{idx}"] = f"=A{idx}+30"  # Data requerido

        # Campos obrigatórios (A→G)
        for col in ["A", "B", "C", "D", "E", "F", "G"]:
            cell = ws[f"{col}{idx}"]
            if not cell.value or str(cell.value).strip() == "":
                cell.fill = red_fill

        # Destaque amarelo em casos de revisão/correção
        if row.get("WasCorrected") or row.get("ValidationStatus") in ("review", "unknown", "no_list"):
            ws[f"D{idx}"].fill = yellow_fill

    # Validação E1:F1
    processed = len(ocr_rows)
    expected  = expected_count
    ws.merge_cells("E1:F1")
    cell = ws["E1"]
    val_str = f"{expected if expected is not None else '?'} / {processed}"
    cell.value = f"Nº Amostras: {val_str}"
    cell.font = bold_center
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.fill = red_fill if (expected is not None and expected != processed) else green_fill
    if expected is not None and expected != processed:
        print(f"⚠️ Diferença de nº de amostras: esperado={expected}, processado={processed}")

    # Origem (G1:J1)
    ws.merge_cells("G1:J1")
    pdf_orig_name = os.path.basename(source_pdf) if source_pdf else "(desconhecida)"
    ws["G1"].value = f"Origem: {pdf_orig_name}"
    ws["G1"].font = Font(italic=True, color="555555")
    ws["G1"].alignment = Alignment(horizontal="left", vertical="center")
    ws["G1"].fill = gray_fill

    # Timestamp (K1:L1)
    ws.merge_cells("K1:L1")
    ws["K1"].value = f"Processado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    ws["K1"].font = Font(italic=True, color="555555")
    ws["K1"].alignment = Alignment(horizontal="right", vertical="center")
    ws["K1"].fill = gray_fill

    base_name = os.path.splitext(os.path.basename(out_name))[0]
    out_path = os.path.join(OUTPUT_DIR, f"{base_name}.xlsx")
    wb.save(out_path)
    print(f"🟢 Gravado (com validação E1/F1, origem G1:J1 e timestamp K1:L1): {out_path}")
    return out_path

def append_process_log(pdf_name, req_id, processed, expected, out_path=None, status="OK", error_msg=None):
    log_path = os.path.join(OUTPUT_DIR, "process_log.csv")
    today_str = datetime.now().strftime("%Y-%m-%d")
    summary_path = os.path.join(OUTPUT_DIR, f"process_summary_{today_str}.txt")

    file_exists = os.path.exists(log_path)
    with open(log_path, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f, delimiter=";")
        if not file_exists:
            writer.writerow([
                "DataHora", "PDF", "ReqID", "Processadas",
                "Requisitadas", "OutputExcel", "Status", "Mensagem"
            ])
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        pdf_short = os.path.basename(pdf_name)
        msg = error_msg or ""
        writer.writerow([timestamp, pdf_short, req_id, processed, expected or "", out_path or "", status, msg])

    if status == "OK":
        print(f"🧾 Log: {os.path.basename(pdf_name)} (req {req_id}) registado como ✅ OK.")
    elif status == "Vazia":
        print(f"🧾 Log: {os.path.basename(pdf_name)} (req {req_id}) registado como ⚠️ Vazia.")
    else:
        print(f"🧾 Log: {os.path.basename(pdf_name)} (req {req_id}) registado como ❌ Erro.")

    try:
        with open(summary_path, "a", encoding="utf-8") as f:
            f.write(f"[{timestamp}] {os.path.basename(pdf_name)} (req {req_id}) → {status} "
                    f"({processed}/{expected or '?'} amostras) {os.path.basename(out_path or '')}\n")
    except Exception as e:
        print(f"⚠️ Falha ao atualizar resumo diário: {e}")

# ───────────────────────────────────────────────────────────────
# 5) PIPELINE: OCR+Parsing para um PDF (assíncrono)
# ───────────────────────────────────────────────────────────────
async def process_pdf_async(pdf_path: str, session: aiohttp.ClientSession):
    print(f"\n📄 A processar async: {os.path.basename(pdf_path)}")
    t0 = asyncio.get_event_loop().time()

    cache_dir = os.path.join(OUTPUT_DIR, "_ocr_cache")
    os.makedirs(cache_dir, exist_ok=True)

    # Render PDF -> imagens (thread pool)
    with ThreadPoolExecutor() as pool:
        images: List[Image.Image] = await asyncio.get_event_loop().run_in_executor(
            pool, lambda: pdf_to_images_pymupdf(pdf_path)
        )

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
        return []

    results.sort(key=lambda x: x[0])
    full_text = "\n".join([f"\n\n--- PÁGINA {i} ---\n{text}" for i, text, _ in results])
    all_tables = [t for _, _, tbls in results for t in tbls]

    t_ocr = asyncio.get_event_loop().time() - t0
    print(f"⏱️ OCR paralelo total: {timedelta(seconds=round(t_ocr))}")

    # Guardar texto global (debug)
    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
    txt_path = os.path.join(OUTPUT_DIR, base_name + "_ocr_debug.txt")
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write(full_text)

    # Parsing
    combined_json = {"analyzeResult": {"tables": all_tables, "pages": []}}
    rows = parse_xylella_from_result(combined_json, pdf_path, txt_path)

    return rows

# ───────────────────────────────────────────────────────────────
# 6) Orquestra parsing multi-requisição (igual ao teu)
# ───────────────────────────────────────────────────────────────
def parse_xylella_from_result(result_json: dict, pdf_name: str, txt_path: Optional[str] = None):
    print(f"📄 A processar: {pdf_name}")

    # Texto global
    if txt_path and os.path.exists(txt_path):
        with open(txt_path, "r", encoding="utf-8") as f:
            full_text = f.read()
        print(f"📝 Contexto extraído de {os.path.basename(txt_path)}")
    else:
        print("⚠️ Ficheiro de texto global não encontrado — fallback sem páginas.")
        full_text = ""

    # Detetar nº de requisições
    count, _ = detect_requisicoes(full_text)

    if count <= 1:
        print("📄 Documento contém apenas uma requisição.")
        context = extract_context_from_text(full_text)
        amostras = parse_xylella_tables(result_json, context, req_id=1)

        if not amostras:
            append_process_log(pdf_name, 1, 0, context.get("declared_samples"),
                               out_path=None, status="Vazia",
                               error_msg="Sem amostras válidas.")
            return []

        out_path = write_to_template(
            amostras,
            os.path.basename(pdf_name),
            expected_count=context.get("declared_samples"),
            source_pdf=pdf_name
        )
        append_process_log(pdf_name, 1, len(amostras),
                           context.get("declared_samples"), out_path=out_path, status="OK")
        return amostras

    # Várias requisições
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

# ───────────────────────────────────────────────────────────────
# 7) Processar TODOS os PDFs da pasta local Input/ (assíncrono)
# ───────────────────────────────────────────────────────────────
async def process_folder_async(input_dir: str = INPUT_DIR):
    pdfs = [os.path.join(input_dir, f) for f in os.listdir(input_dir) if f.lower().endswith(".pdf")]
    if not pdfs:
        print("ℹ️ Não há PDFs na pasta de entrada.")
        return

    start_time = asyncio.get_event_loop().time()
    total_rows = 0

    async with aiohttp.ClientSession() as session:
        tasks = [process_pdf_async(pdf, session) for pdf in pdfs]
        results = await asyncio.gather(*tasks)

    for pdf, res in zip(pdfs, results):
        if not res:
            continue
        total_rows += len(res)

    total_time = asyncio.get_event_loop().time() - start_time
    avg_time = total_time / len(pdfs) if pdfs else 0

    print("\n📊 Resumo Final")
    print("──────────────────────────────")
    print(f"📄 PDFs processados: {len(pdfs)}")
    print(f"🧾 Total de amostras extraídas: {total_rows}")
    print(f"⏱️ Tempo total: {timedelta(seconds=round(total_time))}")
    print(f"⚙️ Tempo médio por PDF: {timedelta(seconds=round(avg_time))}")
    print(f"📂 Saída: {OUTPUT_DIR}")
    print("──────────────────────────────\n")
