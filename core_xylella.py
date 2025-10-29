# core_xylella.py
# -*- coding: utf-8 -*-
"""
Motor de processamento Xylella (usado pela app Streamlit e por orquestradores).
Implementa:
  - process_pdf(pdf_path) -> rows
  - write_to_template(ocr_rows, out_base_path, expected_count=None, source_pdf=None)
"""

from __future__ import annotations
from pathlib import Path
from typing import Any, Optional
import os

# Caminho do template (podes usar o env exportado pelo adaptador)
TEMPLATE_PATH = Path(os.environ.get("TEMPLATE_PATH", Path(__file__).with_name("TEMPLATE_PXF_SGSLABIP1056.xlsx")))

# ───────────────────────────────────────────────────────────────
# 1) PARSE PDF  → devolve 'rows' (a tua estrutura)
# ───────────────────────────────────────────────────────────────
def process_pdf(pdf_path: str) -> Any:
    """
    Recebe o caminho de um PDF e devolve 'rows' que o write_to_template entende.
    SUBSTITUI o bloco TODO pelo teu parser real (OCR/regex/pdfplumber, etc).
    """
    pdf_path = str(pdf_path)
    if not Path(pdf_path).exists():
        raise FileNotFoundError(f"PDF não encontrado: {pdf_path}")

    # TODO: --- INÍCIO do teu parser real ---
    # Exemplo placeholder (remove e cola o teu código):
    # rows = parse_with_regex_or_ocr(pdf_path)
    # return rows
    raise NotImplementedError("Cola aqui a tua implementação real de process_pdf(pdf_path).")
    # TODO: --- FIM do teu parser real ---


# ───────────────────────────────────────────────────────────────
# 2) ESCREVER EXCEL( s ) a partir de 'rows'
# ───────────────────────────────────────────────────────────────
def write_to_template(
    ocr_rows: Any,
    out_base_path: str,
    expected_count: Optional[int] = None,
    source_pdf: Optional[str] = None,
) -> Any:
    """
    Gera 1+ ficheiros Excel.
    - Usa TEMPLATE_PATH como base (cópia) e escreve os dados de 'ocr_rows'.
    - Quando houver várias requisições, grava: <out_base>_req1.xlsx, _req2.xlsx, ...
    - Se 'expected_count' vier preenchido, valida e marca a célula a vermelho quando divergir.
    SUBSTITUI o bloco TODO pela tua escrita real.
    """
    out_base = Path(out_base_path)
    out_base.parent.mkdir(parents=True, exist_ok=True)

    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(f"Template não encontrado: {TEMPLATE_PATH}")

    # TODO: --- INÍCIO da tua escrita real ---
    # Exemplo placeholder (remove e cola o teu código):
    # from openpyxl import load_workbook
    # wb = load_workbook(TEMPLATE_PATH)
    # ws = wb.active
    # ... escrever cabeçalhos, linhas, validações, etc ...
    # out_path = out_base.with_suffix(".xlsx")         # ou _req1.xlsx, _req2.xlsx, ...
    # wb.save(out_path.as_posix())
    # return {"outputs": [out_path.as_posix()]}
    raise NotImplementedError("Cola aqui a tua implementação real de write_to_template(...).")
    # TODO: --- FIM da tua escrita real ---
