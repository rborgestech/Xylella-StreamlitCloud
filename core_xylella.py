# -*- coding: utf-8 -*-
"""
core_xylella.py — motor principal de processamento Xylella
Responsável por:
- OCR automático (Azure → fallback local)
- Extração de texto e parsing de múltiplas requisições
- Escrita dos resultados em Excel (baseado no TEMPLATE)
"""

import os, re, time, io, pdfplumber, requests
from pathlib import Path
from openpyxl import load_workbook, Workbook

# ===============================================================
# 1) OCR & Extração de Texto
# ===============================================================

def extract_text_with_fallback(pdf_path: str) -> str:
    """
    Extrai texto do PDF. Se falhar, tenta OCR Azure (se configurado).
    """
    text = ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    text += "\n" + t
    except Exception as e:
        print(f"⚠️ Erro ao abrir PDF: {e}")

    if text.strip():
        print("🟢 Texto extraído com sucesso (sem OCR).")
        return text

    print("⚠️ Nenhum texto encontrado — a tentar OCR Azure...")

    azure_key = os.getenv("AZURE_KEY")
    azure_endpoint = os.getenv("AZURE_ENDPOINT")

    if azure_key and azure_endpoint:
        try:
            ocr_url = f"{azure_endpoint}/vision/v3.2/read/analyze"
            headers = {"Ocp-Apim-Subscription-Key": azure_key, "Content-Type": "application/pdf"}
            with open(pdf_path, "rb") as f:
                response = requests.post(ocr_url, headers=headers, data=f)
                response.raise_for_status()

            operation_url = response.headers["Operation-Location"]
            for _ in range(30):
                result = requests.get(operation_url, headers=headers).json()
                if result.get("status") == "succeeded":
                    lines = [line["text"] for r in result["analyzeResult"]["readResults"] for line in r["lines"]]
                    print("🟢 Texto obtido por OCR Azure.")
                    return "\n".join(lines)
                time.sleep(1)
        except Exception as e:
            print(f"⚠️ OCR Azure falhou: {e}")

    print("⚠️ OCR Azure indisponível — a tentar OCR local...")
    try:
        from pdf2image import convert_from_path
        import pytesseract
        pages = convert_from_path(pdf_path)
        ocr_text = "\n".join(pytesseract.image_to_string(p) for p in pages)
        if ocr_text.strip():
            print("🟢 Texto obtido por OCR local (Tesseract).")
            return ocr_text
    except Exception as e:
        print(f"⚠️ OCR local falhou: {e}")

    raise RuntimeError(f"Não foi possível extrair texto de {Path(pdf_path).name}")

# ===============================================================
# 2) Parsing simples das amostras e requisições
# ===============================================================

def process_pdf(pdf_path: str):
    """
    Recebe um PDF e devolve lista de linhas [ [campos...], ... ]
    """
    text = extract_text_with_fallback(pdf_path)
    lines = text.splitlines()

    # detectar início de novas requisições (heurística)
    split_idxs = [i for i, l in enumerate(lines) if re.search(r"Requisição|Requisi[cç][aã]o|DGAV", l)]
    split_idxs = split_idxs or [0]
    sections = [lines[split_idxs[i]: split_idxs[i+1]] if i+1 < len(split_idxs) else lines[split_idxs[i]:]
                for i in range(len(split_idxs))]

    rows = []
    for section in sections:
        block = "\n".join(section)
        date_match = re.findall(r"\d{2}/\d{2}/\d{4}", block)
        especie = re.search(r"(Olea europaea|Lavandula|Pelargonium|Rosmarinus|Cistus|Medicago)", block)
        natureza = "Composta" if "Composta" in block else "Simples"
        zona = "Zona Isenta" if "isenta" in block.lower() else "Desconhecida"

        rows.append([
            date_match[0] if len(date_match) > 0 else "",
            date_match[1] if len(date_match) > 1 else "",
            re.search(r"\d{3,4}/\d{4}/[A-Z]+/\d+", block) or re.search(r"\d{3,4}/\d{4}/[A-Z]+", block),
            especie.group(0) if especie else "",
            natureza,
            zona,
            "DGAV" if "DGAV" in block else ""
        ])
    print(f"📊 Extraídas {len(rows)} linhas de amostras.")
    return rows

# ===============================================================
# 3) Escrita no TEMPLATE Excel
# ===============================================================

def write_to_template(ocr_rows, out_base_path, expected_count=None, source_pdf=None):
    """
    Gera 1 ou vários ficheiros Excel a partir do TEMPLATE base.
    """
    template_path = Path(os.getenv("TEMPLATE_PATH", "TEMPLATE_PXF_SGSLABIP1056.xlsx"))
    if not template_path.exists():
        raise FileNotFoundError(f"TEMPLATE não encontrado: {template_path}")

    if not ocr_rows:
        print("⚠️ Nenhuma linha extraída — ficheiro ignorado.")
        return

    if expected_count and expected_count > 1:
        step = len(ocr_rows) // expected_count
        chunks = [ocr_rows[i:i+step] for i in range(0, len(ocr_rows), step)]
    else:
        chunks = [ocr_rows]

    for i, rows in enumerate(chunks, 1):
        out_path = Path(f"{out_base_path}_req{i}.xlsx")
        wb = load_workbook(template_path)
        ws = wb.active
        start_row = 2
        for r, row in enumerate(rows, start=start_row):
            for c, val in enumerate(row, start=1):
                ws.cell(r, c, val)
        wb.save(out_path)
        print(f"🟢 Gravado: {out_path.name}")
