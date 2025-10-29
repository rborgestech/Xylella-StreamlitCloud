# core_xylella.py
# -*- coding: utf-8 -*-
"""
Core real do processador Xylella – versão funcional (com TEMPLATE SGS).

Responsabilidades:
- Extrair texto de PDFs (OCR Azure ou local);
- Detetar e segmentar requisições automaticamente;
- Escrever cada requisição num ficheiro Excel,
  com base no TEMPLATE_PXF_SGSLABIP1056.xlsx,
  mantendo formatações, validações e fórmulas SGS.
"""

from pathlib import Path
from openpyxl import load_workbook
import os, io, re, requests
from PyPDF2 import PdfReader
import pdfplumber
import pytesseract
from pdf2image import convert_from_path
from PIL import Image

# ───────────────────────────────────────────────────────────────
# Funções auxiliares
# ───────────────────────────────────────────────────────────────

def extract_text_with_fallback(pdf_path: str) -> str:
    """Extrai texto do PDF, tentando primeiro texto nativo e depois OCR."""
    pdf_path = Path(pdf_path)
    text = ""

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text += page.extract_text() or ""
    except Exception:
        pass

    # se não há texto nativo, tenta OCR (Azure ou local)
    if not text.strip():
        azure_key = os.environ.get("AZURE_KEY")
        azure_endpoint = os.environ.get("AZURE_ENDPOINT")
        if azure_key and azure_endpoint:
            try:
                text = azure_ocr_extract(pdf_path, azure_key, azure_endpoint)
            except Exception as e:
                print(f"⚠️ Azure OCR falhou ({e}), a tentar OCR local…")
                text = local_ocr_extract(pdf_path)
        else:
            text = local_ocr_extract(pdf_path)

    if not text.strip():
        raise RuntimeError(f"Não foi possível extrair texto de {pdf_path.name}")

    return text


def azure_ocr_extract(pdf_path: Path, key: str, endpoint: str) -> str:
    """Usa o Azure Computer Vision OCR para extrair texto."""
    ocr_url = f"{endpoint}/vision/v3.2/read/analyze"
    headers = {"Ocp-Apim-Subscription-Key": key, "Content-Type": "application/pdf"}

    with open(pdf_path, "rb") as f:
        response = requests.post(ocr_url, headers=headers, data=f)
    response.raise_for_status()

    # Obter URL de operação
    operation_url = response.headers["Operation-Location"]
    import time
    while True:
        result = requests.get(operation_url, headers={"Ocp-Apim-Subscription-Key": key}).json()
        if result.get("status") in ["succeeded", "failed"]:
            break
        time.sleep(1)

    lines = []
    for r in result.get("analyzeResult", {}).get("readResults", []):
        for l in r.get("lines", []):
            lines.append(l["text"])
    return "\n".join(lines)


def local_ocr_extract(pdf_path: Path) -> str:
    """Fallback OCR local (pytesseract)."""
    images = convert_from_path(pdf_path)
    text = ""
    for img in images:
        text += pytesseract.image_to_string(img, lang="por")
    return text


# ───────────────────────────────────────────────────────────────
# PARSER DAS REQUISIÇÕES
# ───────────────────────────────────────────────────────────────

def parse_requisicoes(text: str):
    """Deteta blocos de requisições no texto e extrai dados estruturados."""
    blocos = re.split(r"(?=\bData da Colheita\b)", text)
    rows = []
    for bloco in blocos:
        if not bloco.strip():
            continue
        data_rec = re.search(r"Data.?Rece[cç][aã]o[:\s]+([\d/]+)", bloco)
        data_col = re.search(r"Data.?Colheita[:\s]+([\d/]+)", bloco)
        codigo = re.search(r"([A-Z]?\d{3,4}/\d{4}/[A-Z]{2,3}/?\d?)", bloco)
        especie = re.search(r"Olea europaea|Cistus albidus|Pelargonium|Lavandula|Rosmarinus|Medicago", bloco, re.I)
        natureza = re.search(r"Simples|Composta", bloco, re.I)
        zona = re.search(r"Zona\s+[A-Za-z]+", bloco)
        responsavel = re.search(r"DGAV|INSA|INIAV|Outros", bloco)
        rows.append([
            data_rec.group(1) if data_rec else "",
            data_col.group(1) if data_col else "",
            codigo.group(1) if codigo else "",
            especie.group(0) if especie else "",
            natureza.group(0) if natureza else "",
            zona.group(0) if zona else "",
            responsavel.group(0) if responsavel else "",
        ])
    return rows


# ───────────────────────────────────────────────────────────────
# FUNÇÃO PRINCIPAL
# ───────────────────────────────────────────────────────────────

def process_pdf(pdf_path: str):
    """Pipeline completo: OCR + parsing."""
    text = extract_text_with_fallback(pdf_path)
    rows = parse_requisicoes(text)
    return rows


# ───────────────────────────────────────────────────────────────
# GERAR EXCEL COM TEMPLATE
# ───────────────────────────────────────────────────────────────

def write_to_template(rows, out_base_path, expected_count=None, source_pdf=None):
    """Grava as requisições num Excel baseado no TEMPLATE original."""
    template_path = os.environ.get("TEMPLATE_PATH")
    if not template_path or not Path(template_path).exists():
        raise FileNotFoundError(f"TEMPLATE não encontrado em {template_path}")

    wb = load_workbook(template_path)
    ws = wb.active  # normalmente “Amostras”

    start_row = 6  # a primeira linha de dados no teu template

    for i, row in enumerate(rows, start=start_row):
        for j, value in enumerate(row, start=1):
            ws.cell(row=i, column=j).value = value

    out_path = f"{out_base_path}_req1.xlsx"
    wb.save(out_path)
    print(f"🟢 Gravado: {out_path}")
    return out_path
