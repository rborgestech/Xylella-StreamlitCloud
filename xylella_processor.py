# -*- coding: utf-8 -*-
"""
xylella_processor.py — camada intermédia entre Streamlit (app.py) e core_xylella.py

Funções expostas:
  • process_pdf(pdf_path) → devolve lista de ficheiros Excel gerados (.xlsx)
  • build_zip(file_paths) → constrói ZIP em memória com .xlsx e logs
"""

import os, io, zipfile, importlib
from pathlib import Path
from typing import List, Dict, Any

# Import dinâmico do core
_CORE_MODULE_NAME = "core_xylella"
core = importlib.import_module(_CORE_MODULE_NAME)

# Diretório de saída
OUTPUT_DIR = Path(os.environ.get("OUTPUT_DIR", Path(__file__).parent / "Output"))
OUTPUT_DIR.mkdir(exist_ok=True)

# ───────────────────────────────────────────────
# Processar PDF
# ───────────────────────────────────────────────
def process_pdf(pdf_path: str) -> List[str]:
    """
    Processa um PDF via core_xylella.
    Cria 1 ficheiro Excel por requisição e devolve a lista dos caminhos.
    """
    print(f"\n📄 A processar: {os.path.basename(pdf_path)}")

    # Chamada ao core — devolve lista [{rows, expected}]
    req_results = core.process_pdf_sync(pdf_path)
    created_files = []

    if not req_results:
        print(f"⚠️ Nenhuma requisição extraída de {os.path.basename(pdf_path)}.")
        return []

    for i, req in enumerate(req_results, start=1):
        rows = req.get("rows", [])
        expected = req.get("expected", 0)

        if not rows:
            print(f"⚠️ Requisição {i}: sem amostras válidas.")
            continue

        base = os.path.splitext(os.path.basename(pdf_path))[0]
        out_name = f"{base}_req{i}.xlsx" if len(req_results) > 1 else f"{base}.xlsx"

        # Gera o ficheiro Excel no diretório configurado
        out_path = core.write_to_template(rows, out_name, expected_count=expected, source_pdf=pdf_path)
        if out_path:
            created_files.append(out_path)

        # Log local
        diff = len(rows) - (expected or 0)
        if expected and diff != 0:
            print(f"⚠️ Requisição {i}: {len(rows)} amostras vs {expected} esperadas (diferença {diff:+d}).")
        else:
            print(f"✅ Requisição {i}: {len(rows)} amostras → {os.path.basename(out_path)}")

    print(f"🏁 {os.path.basename(pdf_path)}: {len(created_files)} ficheiro(s) Excel criados.")
    return created_files


# ───────────────────────────────────────────────
# Gerar ZIP com resultados e logs
# ───────────────────────────────────────────────
def build_zip(file_paths: List[str]) -> bytes:
    """
    Constrói um ZIP em memória com todos os ficheiros válidos (.xlsx + txt + logs).
    Inclui automaticamente os _ocr_debug.txt e logs se existirem no OUTPUT_DIR.
    """
    mem = io.BytesIO()
    with zipfile.ZipFile(mem, "w", compression=zipfile.ZIP_DEFLATED) as z:
        # Incluir ficheiros gerados (.xlsx)
        for p in file_paths:
            if p and os.path.exists(p):
                z.write(p, arcname=os.path.basename(p))

        # Incluir ficheiros auxiliares (txt e logs)
        for extra in OUTPUT_DIR.glob("*_ocr_debug.txt"):
            z.write(extra, arcname=os.path.basename(extra))
        for logf in OUTPUT_DIR.glob("process_log.csv"):
            z.write(logf, arcname=os.path.basename(logf))
        for summ in OUTPUT_DIR.glob("process_summary_*.txt"):
            z.write(summ, arcname=os.path.basename(summ))

    mem.seek(0)
    print(f"📦 ZIP criado com {len(file_paths)} ficheiro(s) Excel e logs.")
    return mem.read()
