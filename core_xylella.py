# core_xylella.py
# -*- coding: utf-8 -*-
"""
Versão funcional simplificada para o Streamlit Cloud.
- Lê PDF (texto ou OCR opcional)
- Deteta requisições automaticamente
- Cria 1 Excel por requisição com base no TEMPLATE
"""

from __future__ import annotations
import os, re, pdfplumber
from pathlib import Path
from typing import Any, Optional
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# Caminho do template (ao lado deste ficheiro)
TEMPLATE_PATH = Path(__file__).with_name("TEMPLATE_PXF_SGSLABIP1056.xlsx")


# -------------------------------------------------------------------
# Extrair texto (sem OCR por enquanto)
# -------------------------------------------------------------------
def extract_text(pdf_path: Path) -> str:
    """Extrai texto pesquisável do PDF com pdfplumber."""
    text_parts = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for p in pdf.pages:
                t = p.extract_text() or ""
                if t.strip():
                    text_parts.append(t)
    except Exception as e:
        print(f"⚠️ Falha ao abrir {pdf_path.name}: {e}")
    return "\n".join(text_parts).strip()


# -------------------------------------------------------------------
# Dividir em requisições (padrão simples e robusto)
# -------------------------------------------------------------------
def split_into_requisicoes(text: str) -> list[str]:
    """
    Divide o texto completo em blocos distintos com base em cabeçalhos de requisição.
    Exemplo: "ReqX", "Requisição nº", "DGAV PROGRAMA DE PROSPEÇÃO"
    """
    if not text:
        return []

    # Padrões típicos — ajusta se precisares
    pattern = re.compile(r"(ReqX\\d+|Requ[ií]si[çc][aã]o\\s*n[º°]|DGAV\\s+PROGRAMA)", re.IGNORECASE)
    indices = [m.start() for m in pattern.finditer(text)]

    if not indices:
        return [text]  # só uma requisição

    blocks = []
    for i, pos in enumerate(indices):
        end = indices[i + 1] if i + 1 < len(indices) else len(text)
        blocks.append(text[pos:end].strip())
    return blocks


# -------------------------------------------------------------------
# Função principal: process_pdf
# -------------------------------------------------------------------
def process_pdf(pdf_path: str) -> list[dict[str, str]]:
    """
    Lê o PDF e devolve lista de blocos [{'index': i, 'text': '...'}].
    """
    pdf_path = Path(pdf_path)
    if not pdf_path.exists():
        raise FileNotFoundError(pdf_path)

    text = extract_text(pdf_path)
    if not text:
        raise RuntimeError(f"Não foi possível extrair texto de {pdf_path.name}")

    blocks = split_into_requisicoes(text)
    return [{"index": i + 1, "text": b} for i, b in enumerate(blocks)]


# -------------------------------------------------------------------
# Escrever 1 Excel por requisição
# -------------------------------------------------------------------
def write_to_template(
    ocr_rows: list[dict[str, str]],
    out_base_path: str,
    expected_count: Optional[int] = None,
    source_pdf: Optional[str] = None,
) -> dict:
    """
    Cria 1 ficheiro Excel por requisição detetada.
    Usa o TEMPLATE base e escreve texto simples (para validação).
    """
    base = Path(out_base_path)
    base.parent.mkdir(parents=True, exist_ok=True)

    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(f"TEMPLATE não encontrado: {TEMPLATE_PATH}")

    outputs = []
    for block in ocr_rows:
        i = block["index"]
        text = block["text"]

        wb = load_workbook(TEMPLATE_PATH)
        ws = wb.active

        ws["A1"] = f"Requisição {i}"
        ws["A3"] = f"Origem: {source_pdf or base.name}"
        ws["A5"] = text[:3000]

        out_path = base.with_name(f"{base.name}_req{i}.xlsx")
        wb.save(out_path.as_posix())
        outputs.append(out_path.as_posix())

    # Se esperado ≠ detetado → sinaliza em vermelho
    if expected_count and expected_count != len(outputs):
        wb = load_workbook(outputs[-1])
        ws = wb.active
        ws["E1"] = "Divergência"
        ws["F1"] = f"Esperado {expected_count}, detetado {len(outputs)}"
        ws["F1"].fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        wb.save(outputs[-1])

    return {"outputs": outputs}
