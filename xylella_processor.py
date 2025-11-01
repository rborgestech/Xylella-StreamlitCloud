# -*- coding: utf-8 -*-
from pathlib import Path
import traceback
from core_xylella import process_pdf_sync  # ⚠️ Usa o teu módulo que fazia o parsing certo
from zipfile import ZipFile
import io


# ───────────────────────────────────────────────
# 1. Função principal
# ───────────────────────────────────────────────
def process_pdf(pdf_path):
    """
    Processa um PDF e gera um ou vários ficheiros Excel.
    Retorna uma lista de tuplos (path, n_amostras, discrepancias).
    Compatível com múltiplas requisições (req1, req2, ...).
    """
    try:
        result = process_pdf_sync(pdf_path)

        # Caso a função original devolva apenas paths:
        if isinstance(result, list) and all(isinstance(x, str) for x in result):
            return [(str(Path(x).resolve()), None, None) for x in result]

        # Caso devolva tuplos (path, amostras, discrepâncias)
        elif isinstance(result, list) and all(isinstance(x, tuple) for x in result):
            normalized = []
            for fp, n, d in result:
                normalized.append((str(Path(fp).resolve()), n, d))
            return normalized

        # Caso devolva dicionários
        elif isinstance(result, list) and all(isinstance(x, dict) for x in result):
            normalized = []
            for r in result:
                normalized.append((
                    str(Path(r.get("path")).resolve()),
                    r.get("samples"),
                    r.get("discrepancies"),
                ))
            return normalized

        # Caso raro — um único ficheiro
        elif isinstance(result, str):
            return [(str(Path(result).resolve()), None, None)]

        else:
            print(f"⚠️ Formato inesperado em process_pdf: {type(result)} → {result}")
            return []

    except Exception as e:
        print("❌ ERRO no process_pdf:", e)
        traceback.print_exc()
        return []


# ───────────────────────────────────────────────
# 2. Função auxiliar para criar ZIP
# ───────────────────────────────────────────────
def build_zip(file_paths):
    """Cria um ZIP em memória com os ficheiros Excel."""
    zip_buffer = io.BytesIO()
    with ZipFile(zip_buffer, "w") as zip_file:
        for fp in file_paths:
            try:
                zip_file.write(fp, arcname=Path(fp).name)
            except Exception as e:
                print(f"⚠️ Erro a adicionar {fp} ao ZIP: {e}")
    zip_buffer.seek(0)
    return zip_buffer.getvalue()
