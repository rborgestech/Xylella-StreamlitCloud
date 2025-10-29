# core_xylella.py
# -*- coding: utf-8 -*-
"""
Core Xylella Processor – versão final (template SGS + split automático + OCR híbrido)
Autor: Rosa Borges

Funções:
 - Extrai texto (OCR Azure/local)
 - Deteta múltiplas requisições por PDF
 - Escreve resultados no TEMPLATE_PXF_SGSLABIP1056.xlsx
   preservando fórmulas, validações e formatação SGS
"""

import os, re, io, time, shutil, requests
from pathlib import Path
from openpyxl import load_workbook
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
#  PARSER – deteção de requisições
# ───────────────────────────────────────────────────────────────

def parse_with_regex(text: str):
    """Extrai blocos de amostras e campos relevantes usando regex."""
    padrao = re.compile(
        r"(?P<data_rec>\d{2}/\d{2}/\d{4}).*?"
        r"(?P<data_col>\d{2}/\d{2}/\d{4}).*?"
        r"(?P<codigo>\d{3,}\/\d{4}\/[A-Z]{2,}|[0-9]{5,})?.*?"
        r"(?P<especie>[A-Z][a-zç]+(?: [a-z]+){0,2}).*?"
        r"(?P<natureza>Simples|Composta).*?"
        r"(?P<zona>Isenta|Contida|Desconhec[ia]do|Zona [A-Za-z]+)?.*?"
        r"(?P<responsavel>DGAV|INIAV|INSA|Outros)?",
        re.S,
    )

    resultados = []
    for m in padrao.finditer(text):
        resultados.append([
            m.group("data_rec") or "",
            m.group("data_col") or "",
            m.group("codigo") or "",
            m.group("especie") or "",
            m.group("natureza") or "",
            m.group("zona") or "",
            m.group("responsavel") or ""
        ])
    return resultados


# ───────────────────────────────────────────────────────────────
#  SPLIT – múltiplas requisições no mesmo PDF
# ───────────────────────────────────────────────────────────────

def split_if_multiple_requisicoes(text: str):
    """Divide o PDF em várias requisições (por cabeçalho de 'Data da Colheita')."""
    indices = [m.start() for m in re.finditer(r"Data.?Colheita", text)]
    if len(indices) <= 1:
        return [text]

    partes = []
    for i in range(len(indices)):
        start = indices[i]
        end = indices[i + 1] if i + 1 < len(indices) else len(text)
        partes.append(text[start:end])
    return partes


# ───────────────────────────────────────────────────────────────
#  PROCESSAMENTO COMPLETO
# ───────────────────────────────────────────────────────────────

def process_pdf(pdf_path: str):
    """Extrai o texto, divide em requisições e devolve listas de linhas."""
    text = extract_text_with_fallback(pdf_path)
    blocos = split_if_multiple_requisicoes(text)
    todas = []
    for bloco in blocos:
        linhas = parse_with_regex(bloco)
        if linhas:
            todas.append(linhas)
    return todas


# ───────────────────────────────────────────────────────────────
#  ESCREVER NO TEMPLATE SGS
# ───────────────────────────────────────────────────────────────

def write_to_template(ocr_rows, out_base_path, expected_count=None, source_pdf=None):
    """
    Escreve as requisições no TEMPLATE_PXF_SGSLABIP1056.xlsx
    mantendo fórmulas, validações e formatação SGS.
    """
    template_path = Path(os.environ["TEMPLATE_PATH"])
    if not template_path.exists():
        raise FileNotFoundError(f"TEMPLATE não encontrado: {template_path}")

    out_files = []
    start_row = 6
    sheet_name = "Avaliação pré registo"

    for idx, req_rows in enumerate(ocr_rows, start=1):
        out_path = Path(f"{out_base_path}_req{idx}.xlsx")
        shutil.copy(template_path, out_path)

        wb = load_workbook(out_path)
        ws = wb[sheet_name]

        for i, row in enumerate(req_rows, start=start_row):
            for j, value in enumerate(row, start=1):
                ws.cell(row=i, column=j).value = value

        wb.save(out_path)
        print(f"🟢 Gravado com sucesso: {out_path}")
        out_files.append(out_path)

    return out_files
