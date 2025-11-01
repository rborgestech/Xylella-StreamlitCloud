# -*- coding: utf-8 -*-
import os
import shutil
from pathlib import Path
from datetime import datetime

try:
    from core_xylella import process_pdf_sync
except ImportError:
    process_pdf_sync = None


def process_pdf(pdf_path):
    """
    Wrapper estável para processar PDFs via core_xylella.
    - Chama diretamente process_pdf_sync do core.
    - Garante que o core trabalha no diretório do projeto (para criar debug/ e summary).
    - Retorna lista de dicionários: [{path, processed, discrepancy}, ...]
    """
    pdf_path = Path(pdf_path).resolve()
    project_root = Path.cwd()
    pdf_name = pdf_path.name

    # Copia o ficheiro para o diretório do projeto antes de processar
    stable_copy = project_root / pdf_name
    if not stable_copy.exists():
        shutil.copy(pdf_path, stable_copy)

    # Define o diretório de trabalho igual ao do projeto
    os.chdir(project_root)

    if not process_pdf_sync:
        print("⚠️ core_xylella não encontrado — devolve simulação.")
        excel_path = stable_copy.with_suffix(".xlsx")
        return [{"path": str(excel_path), "processed": 0, "discrepancy": False}]

    print(f"🧪 Início de processamento: {stable_copy.name}")
    result = process_pdf_sync(str(stable_copy))

    # Volta ao diretório original (por segurança)
    os.chdir(Path(__file__).parent)

    # Normaliza resultados
    entries = []
    if isinstance(result, list):
        for r in result:
            if isinstance(r, str):
                entries.append({"path": r, "processed": 0, "discrepancy": False})
            elif isinstance(r, dict):
                entries.append(r)
            elif isinstance(r, tuple) and len(r) >= 1:
                entries.append({
                    "path": r[0],
                    "processed": r[1] if len(r) > 1 else 0,
                    "discrepancy": bool(r[2]) if len(r) > 2 else False
                })

    print("✅ Processamento concluído no diretório do projeto.")
    return entries


def build_zip(paths):
    """Gera ZIP a partir de paths (strings ou dicts)."""
    import io, zipfile
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as z:
        for p in paths:
            if isinstance(p, dict):
                p = p.get("path")
            p = Path(p)
            if p.exists():
                z.write(p, arcname=p.name)
    zip_buffer.seek(0)
    return zip_buffer.getvalue()
