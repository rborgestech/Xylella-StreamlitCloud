# -*- coding: utf-8 -*-
"""
Módulo Xylella Processor
Responsável por processar PDFs de requisições e gerar ficheiros Excel por requisição.
Inclui normalização do output para integração com o front-end Streamlit.
"""

import os
from pathlib import Path
from typing import List, Dict, Any
import importlib   # ✅ necessário para carregar o core dinamicamente

# Carrega o módulo core_xylella dinamicamente
core = importlib.import_module("core_xylella")


# ----------------------------------------------------------------------
# Normalizador universal — converte qualquer formato em lista de dicionários
# ----------------------------------------------------------------------
def _as_list_of_entries(result):
    entries = []

    # Caso 1: lista de dicionários ou tuples
    if isinstance(result, list):
        for item in result:
            if isinstance(item, dict):
                p = item.get("path") or item.get("filepath") or item.get("file")
                if not p:
                    continue
                entries.append({
                    "path": str(p),
                    "samples": item.get("samples") or item.get("amostras"),
                    "discrepancy": item.get("discrepancy") or item.get("discrepancias") or 0
                })
            elif isinstance(item, (tuple, list)) and len(item) >= 1:
                p = item[0]
                smp = item[1] if len(item) > 1 else None
                dsc = item[2] if len(item) > 2 else 0
                if p:
                    entries.append({"path": str(p), "samples": smp, "discrepancy": dsc})

    # Caso 2: tuplo com mapas ou agregados
    elif isinstance(result, tuple) and len(result) >= 1:
        paths = result[0] or []
        if len(result) >= 3 and isinstance(result[1], dict) and isinstance(result[2], dict):
            samples_map = result[1]
            disc_map = result[2]
            for p in paths:
                if not p:
                    continue
                entries.append({
                    "path": str(p),
                    "samples": samples_map.get(p),
                    "discrepancy": disc_map.get(p, 0)
                })
        else:
            for p in paths:
                if not p:
                    continue
                entries.append({"path": str(p), "samples": None, "discrepancy": 0})

    return entries


# ----------------------------------------------------------------------
# Função pública usada pelo Streamlit
# ----------------------------------------------------------------------
def process_pdf(pdf_path: str) -> List[str]:
    """
    Processa um PDF via core_xylella e devolve a lista de caminhos .xlsx criados.
    Aguenta 3 formatos de retorno do core:
      A) List[Dict{"rows": [...], "declared": n}]  -> escreve 1 xlsx por req
      B) List[List[Dict]]                          -> escreve 1 xlsx por req
      C) List[str]                                 -> já são caminhos xlsx -> devolve tal como estão
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

    def _write_one_req(rows: list, expected: int | None, req_idx: int, total_reqs: int):
        """Escreve uma requisição (lista de dicts) no template e retorna o caminho."""
        if not rows or not isinstance(rows, list):
            return None
        if not all(isinstance(r, dict) for r in rows):
            print(f"⚠️ Req {req_idx}: formato inesperado (não é lista de dicts). Ignorado.")
            return None

        out_name = f"{base}_req{req_idx}.xlsx" if total_reqs > 1 else f"{base}.xlsx"
        out_path = core.write_to_template(
            rows,
            out_name,
            expected_count=expected,
            source_pdf=pdf_path
        )
        if out_path and os.path.exists(out_path):
            print(f"✅ Requisição {req_idx}: {len(rows)} amostras → {os.path.basename(out_path)}")
            return out_path
        return None

    # Caso A) várias requisições — formato [{rows, declared}]
    if isinstance(req_results, list) and all(isinstance(x, dict) for x in req_results):
        total = len(req_results)
        for i, req in enumerate(req_results, start=1):
            rows = req.get("rows", [])
            expected = req.get("declared") or req.get("expected") or None
            p = _write_one_req(rows, expected, i, total)
            if p:
                created_files.append(p)
        print(f"🏁 {base}: {len(created_files)} ficheiro(s) Excel criados.")
        return created_files

    # Caso B) lista de listas (formato antigo)
    if isinstance(req_results, list) and all(isinstance(x, list) for x in req_results):
        total = len(req_results)
        for i, rows in enumerate(req_results, start=1):
            p = _write_one_req(rows, None, i, total)
            if p:
                created_files.append(p)
        print(f"🏁 {base}: {len(created_files)} ficheiro(s) Excel criados.")
        return created_files

    # Formato desconhecido
    print(f"⚠️ Formato de retorno inesperado de core.process_pdf_sync para {base}.")
    return []


# ----------------------------------------------------------------------
# Função auxiliar de ZIP (mantém a tua implementação)
# ----------------------------------------------------------------------
def build_zip(paths_or_entries):
    """Cria ZIP a partir de paths ou lista de entries."""
    import io, zipfile

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for item in paths_or_entries:
            p = item["path"] if isinstance(item, dict) else item
            if os.path.exists(p):
                zf.write(p, arcname=os.path.basename(p))
    zip_buffer.seek(0)
    return zip_buffer.getvalue()
