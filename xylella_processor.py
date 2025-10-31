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
    Processa um PDF via core_xylella e devolve a lista de caminhos .xlsx criados.
    Aguenta 3 formatos de retorno do core:
      A) List[List[Dict]]  -> escreve 1 xlsx por req
      B) List[Dict]        -> escreve 1 xlsx
      C) List[str]         -> já são caminhos xlsx -> devolve tal como estão
    """
    print(f"\n📄 A processar: {os.path.basename(pdf_path)}")
    base = os.path.splitext(os.path.basename(pdf_path))[0]

    req_results = core.process_pdf_sync(pdf_path)
    if not req_results:
        print(f"⚠️ Nenhuma requisição extraída de {base}.")
        return []

    # Caso C) já são ficheiros .xlsx (strings)
    if isinstance(req_results, list) and all(isinstance(x, str) for x in req_results):
        created_files = [p for p in req_results if os.path.exists(p)]
        print(f"🟢 Core devolveu {len(created_files)} ficheiros já criados.")
        return created_files

    created_files: List[str] = []

    def _write_one_req(rows: list, req_idx: int, total_reqs: int):
        """Escreve uma requisição (lista de dicts) no template e retorna o caminho."""
        if not rows or not isinstance(rows, list):
            return None
        if not all(isinstance(r, dict) for r in rows):
            # proteção extra: se por algum motivo vierem strings aqui, ignora
            print(f"⚠️ Req {req_idx}: formato inesperado (não é lista de dicts). Ignorado.")
            return None

        # tenta obter expected se vier embutido em cada row (compatibilidade futura)
        expected = None
        try:
            if rows and isinstance(rows[0], dict) and "expected" in rows[0]:
                expected = rows[0].get("expected")
        except Exception:
            expected = None

        out_name = f"{base}_req{req_idx}.xlsx" if total_reqs > 1 else f"{base}.xlsx"
        out_path = core.write_to_template(rows, out_name, expected_count=expected, source_pdf=pdf_path)
        if out_path and os.path.exists(out_path):
            print(f"✅ Requisição {req_idx}: {len(rows)} amostras → {os.path.basename(out_path)}")
            return out_path
        return None

    # Caso B) uma única requisição (lista de dicts)
    if isinstance(req_results, list) and req_results and all(isinstance(x, dict) for x in req_results):
        p = _write_one_req(req_results, 1, 1)
        return [p] if p else []

    # Caso A) várias requisições (lista de listas de dicts)
    if isinstance(req_results, list) and all(isinstance(x, list) for x in req_results):
        total = len(req_results)
        for i, rows in enumerate(req_results, start=1):
            p = _write_one_req(rows, i, total)
            if p:
                created_files.append(p)
        print(f"🏁 {base}: {len(created_files)} ficheiro(s) Excel criados.")
        return created_files

    # Formato desconhecido — não faz nada
    print(f"⚠️ Formato de retorno inesperado de core.process_pdf_sync para {base}.")
    return []


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
