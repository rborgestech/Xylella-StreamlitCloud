# -*- coding: utf-8 -*-
"""
xylella_processor.py — versão final estável
"""
import io
import traceback
from pathlib import Path
from zipfile import ZipFile
from core_xylella import process_pdf_sync  # <-- garante que chama o teu parser real


# ───────────────────────────────────────────────
# Função principal
# ───────────────────────────────────────────────
def process_pdf(pdf_path):
    """
    Processa um PDF e gera um ou mais ficheiros Excel.
    Retorna lista de tuplos (path, solicitadas, processadas)
    """
    try:
        # Cria pasta de debug
        DEBUG_DIR = Path("debug")
        DEBUG_DIR.mkdir(exist_ok=True)

        result = process_pdf_sync(pdf_path)
        normalized = []

        if not result:
            print(f"⚠️ Nenhum resultado devolvido para {pdf_path}")
            return []

        for item in result:
            # Suporta dicts ou tuplos
            if isinstance(item, dict):
                fp = item.get("path")
                solicitadas = item.get("samples_requested") or item.get("samples") or 0
                processadas = item.get("samples_processed") or item.get("processed") or 0
            elif isinstance(item, tuple):
                fp, solicitadas, processadas = item + (0,) * (3 - len(item))
            else:
                fp, solicitadas, processadas = str(item), 0, 0

            normalized.append((str(Path(fp).resolve()), solicitadas, processadas))

        # Log detalhado na pasta debug
        log_path = DEBUG_DIR / f"{Path(pdf_path).stem}_debug.log"
        with open(log_path, "w", encoding="utf-8") as logf:
            logf.write(f"📄 {Path(pdf_path).name}\n\n")
            for fp, s, p in normalized:
                logf.write(f"{fp} | solicitadas={s} | processadas={p}\n")

        return normalized

    except Exception as e:
        print("❌ ERRO no process_pdf:", e)
        traceback.print_exc()
        return []


# ───────────────────────────────────────────────
# Criação do ZIP
# ───────────────────────────────────────────────
def build_zip(file_paths):
    zip_buffer = io.BytesIO()
    with ZipFile(zip_buffer, "w") as zip_file:
        for fp in file_paths:
            try:
                zip_file.write(fp, arcname=Path(fp).name)
            except Exception as e:
                print(f"⚠️ Erro a adicionar {fp} ao ZIP: {e}")
    zip_buffer.seek(0)
    return zip_buffer.getvalue()
