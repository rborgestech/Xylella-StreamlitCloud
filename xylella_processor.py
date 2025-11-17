# -*- coding: utf-8 -*-
"""
Módulo Xylella Processor
Encapsula apenas a chamada ao core_xylella.
"""

import os
from pathlib import Path
import importlib

# Carregar o core
core = importlib.import_module("core_xylella")

def process_pdf(pdf_path: str):
    """
    Processa um PDF via core_xylella e devolve a lista de caminhos .xlsx criados.
    O core já devolve exatamente isso → List[str]
    """
    print(f"\n📄 A processar: {os.path.basename(pdf_path)}")

    created_files = core.process_pdf_sync(pdf_path)

    # Garantir que são paths válidos
    created_files = [p for p in created_files if p and Path(p).exists()]

    print(f"🟢 {len(created_files)} ficheiro(s) Excel criados.")
    return created_files


def build_zip(paths):
    """Cria ZIP a partir de paths válidos."""
    import io, zipfile

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for p in paths:
            if Path(p).exists():
                zf.write(p, arcname=Path(p).name)
    zip_buffer.seek(0)
    return zip_buffer.getvalue()
