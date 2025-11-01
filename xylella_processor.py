# -*- coding: utf-8 -*-
import os
from pathlib import Path
from datetime import datetime

# Import seguro do core real
try:
    from core_xylella import process_pdf_sync
except ImportError:
    process_pdf_sync = None


def process_pdf(pdf_path):
    """
    Wrapper que invoca o processador real (core_xylella).
    Cria automaticamente pasta debug/ e summary.
    """
    if not process_pdf_sync:
        print("⚠️ core_xylella não encontrado — devolve lista simulada.")
        excel_path = Path(pdf_path).with_suffix(".xlsx")
        return [{"path": str(excel_path), "processed": 0, "discrepancy": False}]

    # ⚠️ Corrige caminho para absoluto
    pdf_path = Path(pdf_path).resolve()
    if not pdf_path.exists():
        raise FileNotFoundError(f"Ficheiro não encontrado: {pdf_path}")

    pdf_name = pdf_path.stem
    debug_dir = Path.cwd() / "debug"
    debug_dir.mkdir(exist_ok=True)

    print(f"🧪 Início de processamento: {pdf_path.name}")
    result = process_pdf_sync(str(pdf_path))

def build_zip(paths):
    """
    Gera um ZIP com os paths fornecidos.
    """
    import io, zipfile
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as z:
        for p in paths:
            p = Path(p)
            if p.exists():
                z.write(p, arcname=p.name)
    zip_buffer.seek(0)
    return zip_buffer.getvalue()
