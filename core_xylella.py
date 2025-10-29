# core_xylella.py
# -*- coding: utf-8 -*-
"""
Core Xylella Processor – versão funcional com TEMPLATE SGS

Funções:
 - Extrai texto (OCR Azure/local)
 - Deteta múltiplas requisições por PDF
 - Escreve resultados em cópias do TEMPLATE_PXF_SGSLABIP1056.xlsx
   preservando fórmulas, validações e formatação
"""

import os, re, io, time, shutil, requests
from pathlib import Path
from openpyxl import load_workbook
from PyPDF2 import PdfReader
from pdf2image import convert_from_path
from PIL import Image
import pytesseract
import pdfplumber

# ───────────────────────────────────────────────────────────────
#  OCR – Azure ou local
# ───────────────────────────────────────────────────────────────

def extract_text_with_fallback(pdf_path: str) -> str:
    """Extrai texto de um PDF via texto nativo, Azure OCR ou Tesseract local."""
    pdf_path = Path(pdf_path)
    text = ""

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text += page.extract_text() or ""
    except Exception:
        pass

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
    """Usa Azure Computer Vision (OCR)"""
    ocr_url = f"{endpoint}/vision/v3.2/read/analyze"
    headers = {"Ocp-Apim-Subscription-Key": key, "Content-Type": "application/pdf"}

    with open(pdf_path, "rb") as f:
        response = requests.post(ocr_url, headers=headers, data=f)
    response.raise_for_status()

    operation_url = response.headers["Operation-Location"]
    while True:
        result = requests.get(operation_url, headers=headers).json()
        if result.get("status") in ["succeeded", "failed"]:
            break
        time.sleep(1)

    lines = []
    for r in result.get("analyzeResult", {}).get("readResults", []):
        for l in r.get("lines", []):
            lines.append(l["text"])
    return "\n".join(lines)


def local_ocr_extract(pdf_path: Path) -> str:
    """OCR local com Tesseract"""
    images = convert_from_path(pdf_path)
    text = ""
    for img in images:
        text += pytesseract.image_to_string(img, lang="por")
    return text


# ───────────────────────────────────────────────────────────────
#  PARSER – deteta blocos de requisições
# ───────────────────────────────────────────────────────────────

def parse_requisicoes(text: str):
    """Identifica blocos de requisições e extrai dados estruturados."""
    blocos = re.split(r"(?=\bData da Colheita\b)", text)
    all_reqs = []

    for bloco in blocos:
        if not bloco.strip():
            continue

        data_rec = re.search(r"Data.?Rece[cç][aã]o[:\s]+([\d/]+)", bloco)
        data_col = re.search(r"Data.?Colheita[:\s]+([\d/]+)", bloco)
        codigo = re.search(r"([A-Z]?\d{3,4}/\d{4}/[A-Z]{2,3}/?\d?)", bloco)
        especie = re.search(r"Olea europaea|Cistus albidus|Pelargonium|Lavandula|Rosmarinus|Medicago", bloco, re.I)
        natureza = re.search(r"Simples|Composta", bloco, re.I)
        zona = re.search(r"Zona\s+[A-Za-z]+", bloco)
        responsavel = re.search(r"DGAV|INIAV|INSA|Outros", bloco)

        row = [
            data_rec.group(1) if data_rec else "",
            data_col.group(1) if data_col else "",
            codigo.group(1) if codigo else "",
            especie.group(0) if especie else "",
            natureza.group(0) if natureza else "",
            zona.group(0) if zona else "",
            responsavel.group(0) if responsavel else "",
        ]
        all_reqs.append(row)

    return all_reqs


# ───────────────────────────────────────────────────────────────
#  PROCESSAMENTO PRINCIPAL
# ───────────────────────────────────────────────────────────────

def process_pdf(pdf_path: str):
    """Executa OCR + parsing completo."""
    text = extract_text_with_fallback(pdf_path)
    rows = parse_requisicoes(text)
    return rows


# ───────────────────────────────────────────────────────────────
#  ESCREVER NO TEMPLATE SGS
# ───────────────────────────────────────────────────────────────

def write_to_template(ocr_rows, out_base_path, expected_count=None, source_pdf=None):
    """
    Grava as requisições no TEMPLATE_PXF_SGSLABIP1056.xlsx,
    mantendo fórmulas, validações e formato SGS.
    """
    template_path = Path(os.environ["TEMPLATE_PATH"])
    if not template_path.exists():
        raise FileNotFoundError(f"TEMPLATE não encontrado: {template_path}")

    out_files = []
    start_row = 6  # linha onde começam as amostras
    sheet_name = "Amostras"  # nome da folha do template

    # Cada bloco corresponde a uma requisição → ficheiro novo
    for idx, rowset in enumerate(split_requisicoes(ocr_rows), start=1):
        out_path = Path(f"{out_base_path}_req{idx}.xlsx")
        shutil.copy(template_path, out_path)

        wb = load_workbook(out_path)
        ws = wb[sheet_name]

        for i, row in enumerate(rowset, start=start_row):
            for j, value in enumerate(row, start=1):
                ws.cell(row=i, column=j).value = value

        wb.save(out_path)
        print(f"🟢 Gravado: {out_path}")
        out_files.append(out_path)

    return out_files


# ───────────────────────────────────────────────────────────────
#  SUPORTE – dividir blocos de requisições
# ───────────────────────────────────────────────────────────────

def split_requisicoes(rows):
    """Divide o conjunto de linhas em blocos (1 por requisição)."""
    if not rows:
        return []
    # Heurística: cada requisição contém até 50 linhas no máximo
    step = 50 if len(rows) > 50 else len(rows)
    return [rows[i:i + step] for i in range(0, len(rows), step)]
