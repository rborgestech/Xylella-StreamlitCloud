# core_xylella.py
# -*- coding: utf-8 -*-
"""
Core Xylella – Processamento real (deteta requisições e usa OCR se necessário).
Autor: Rosa Borges
Data: 2025-10-30
"""

from __future__ import annotations
import os, re, io, time, requests, pdfplumber
from pathlib import Path
from typing import Any, Optional
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# Caminho do TEMPLATE
TEMPLATE_PATH = Path(os.environ.get("TEMPLATE_PATH", Path(__file__).with_name("TEMPLATE_PXF_SGSLABIP1056.xlsx")))

# -------------------------------------------------------------------
# 🧩 Função principal: process_pdf(pdf_path)
# -------------------------------------------------------------------
def process_pdf(pdf_path: str) -> list[dict[str, str]]:
    """
    Extrai texto de 1 PDF (OCR se necessário) e deteta requisições.
    Devolve lista de blocos: [{'index': 1, 'text': '...'}, ...]
    """
    pdf_path = Path(pdf_path)
    if not pdf_path.exists():
        raise FileNotFoundError(pdf_path)

    text = extract_text_with_fallback(pdf_path)
    if not text.strip():
        raise RuntimeError(f"Nenhum texto extraído de {pdf_path.name}")

    # Divide o texto com base em padrões típicos de cabeçalho de requisição
    blocks = split_into_requisicoes(text)
    return [{"index": i + 1, "text": b} for i, b in enumerate(blocks)]


# -------------------------------------------------------------------
# 🔍 Extração de texto (pdfplumber → Azure OCR se falhar)
# -------------------------------------------------------------------
def extract_text_with_fallback(pdf_path: Path) -> str:
    text_parts = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for p in pdf.pages:
                t = p.extract_text() or ""
                text_parts.append(t)
    except Exception as e:
        print(f"⚠️ Falha ao extrair com pdfplumber: {e}")

    text = "\n".join(text_parts).strip()
    if len(text) > 50:
        return text

    print("📄 Texto insuficiente – a tentar OCR via Azure...")
    azure_key = os.getenv("AZURE_OCR_KEY")
    azure_endpoint = os.getenv("AZURE_OCR_ENDPOINT")
    if not azure_key or not azure_endpoint:
        raise RuntimeError("AZURE_OCR_KEY/AZURE_OCR_ENDPOINT não configurados nos secrets do Streamlit.")

    return azure_ocr_extract(pdf_path, azure_key, azure_endpoint)


# -------------------------------------------------------------------
# 🧠 Azure OCR (Read API v3.2)
# -------------------------------------------------------------------
def azure_ocr_extract(pdf_path: Path, key: str, endpoint: str) -> str:
    ocr_url = f"{endpoint.rstrip('/')}/vision/v3.2/read/analyze"
    headers = {"Ocp-Apim-Subscription-Key": key, "Content-Type": "application/pdf"}
    with open(pdf_path, "rb") as f:
        response = requests.post(ocr_url, headers=headers, data=f)
    if response.status_code not in (200, 202):
        raise RuntimeError(f"OCR request failed ({response.status_code}): {response.text}")
    operation_url = response.headers.get("Operation-Location")
    if not operation_url:
        raise RuntimeError("Azure OCR não devolveu Operation-Location.")

    # polling
    for _ in range(30):
        result = requests.get(operation_url, headers={"Ocp-Apim-Subscription-Key": key}).json()
        status = result.get("status")
        if status == "succeeded":
            lines = []
            for page in result["analyzeResult"]["readResults"]:
                for line in page["lines"]:
                    lines.append(line["text"])
            return "\n".join(lines)
        elif status == "failed":
            raise RuntimeError("Azure OCR falhou a análise.")
        time.sleep(1)
    raise TimeoutError("Azure OCR expirou após 30s.")


# -------------------------------------------------------------------
# ✂️ Split em múltiplas requisições
# -------------------------------------------------------------------
def split_into_requisicoes(text: str) -> list[str]:
    """
    Divide o texto completo em blocos por cabeçalhos típicos de requisição.
    Ajusta os padrões conforme o formato real dos teus PDFs.
    """
    # Exemplo: linhas que contêm códigos tipo "ReqX02", "ReqX03" ou "Requisição nº"
    pattern = re.compile(r"(ReqX\\d+|Requ[ií]si[çc][aã]o\\s*n[º°])", re.IGNORECASE)
    indices = [m.start() for m in pattern.finditer(text)]
    if not indices:
        # se não encontrar separadores, devolve tudo num bloco
        return [text]
    blocks = []
    for i, pos in enumerate(indices):
        end = indices[i + 1] if i + 1 < len(indices) else len(text)
        blocks.append(text[pos:end].strip())
    return blocks


# -------------------------------------------------------------------
# 🧾 Escrever 1 Excel por requisição detectada
# -------------------------------------------------------------------
def write_to_template(
    ocr_rows: list[dict[str, str]],
    out_base_path: str,
    expected_count: Optional[int] = None,
    source_pdf: Optional[str] = None,
) -> dict:
    base = Path(out_base_path)
    base.parent.mkdir(parents=True, exist_ok=True)

    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(f"Template não encontrado: {TEMPLATE_PATH}")

    outputs = []
    for block in ocr_rows:
        i = block["index"]
        text = block["text"]
        wb = load_workbook(TEMPLATE_PATH)
        ws = wb.active

        ws["A1"] = f"Requisição {i}"
        ws["A3"] = f"Origem: {source_pdf or base.name}"
        ws["A5"] = text[:3000]  # grava texto parcial

        out_path = base.with_name(f"{base.name}_req{i}.xlsx")
        wb.save(out_path.as_posix())
        outputs.append(out_path.as_posix())

    # Validação simples
    if expected_count and expected_count != len(outputs):
        wb = load_workbook(outputs[-1])
        ws = wb.active
        ws["E1"] = "Divergência"
        ws["F1"] = f"Esperado {expected_count}, detetado {len(outputs)}"
        ws["F1"].fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        wb.save(outputs[-1])

    return {"outputs": outputs}
