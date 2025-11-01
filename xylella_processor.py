# -*- coding: utf-8 -*-
"""
Módulo Xylella Processor
Responsável por processar PDFs de requisições e gerar ficheiros Excel por requisição.
Inclui normalização do output para integração com o front-end Streamlit.
"""

import os
from pathlib import Path

# ----------------------------------------------------------------------
# Função original — mantém a tua lógica existente aqui
# ----------------------------------------------------------------------
def process_pdf_original(pdf_path):
    """
    Implementação original de processamento.
    Deve devolver um dos seguintes formatos:
      1. [(path, samples, discrepancy), ...]
      2. [{"path": ..., "samples": ..., "discrepancy": ...}, ...]
      3. ([paths], samples_map, discrepancy_map)
      4. ([paths], total_samples, total_discrepancies)
    """
    # ⚠️ Substitui este exemplo pela tua implementação real:
    excel_path = Path(pdf_path).with_suffix(".xlsx")
    # Simulação de resultado
    return [(str(excel_path), 12, 0)]


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
            # ([paths], samples_map, discrepancy_map)
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
            # ([paths], total_samples, total_discrepancies)
            for p in paths:
                if not p:
                    continue
                entries.append({"path": str(p), "samples": None, "discrepancy": 0})

    return entries


# ----------------------------------------------------------------------
# Função pública usada pelo Streamlit
# ----------------------------------------------------------------------
def process_pdf(pdf_path):
    """Wrapper estável — garante lista de dicionários por ficheiro."""
    result = process_pdf_original(pdf_path)
    entries = _as_list_of_entries(result)
    for e in entries:
        e["path"] = str(Path(e["path"]).resolve())
    return entries


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
