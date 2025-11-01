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
    Wrapper estável — força execução no diretório do projeto.
    Garante que debug/ e summary são criados como no teste.
    """
    pdf_path = Path(pdf_path).resolve()
    project_root = Path(__file__).parent.resolve()
    pdf_name = pdf_path.name

    # Copiar o PDF carregado para o diretório do projeto
    stable_copy = project_root / pdf_name
    shutil.copy(pdf_path, stable_copy)
    print(f"📄 Copiado para {stable_copy}")

    # ⚙️ Forçar diretório de trabalho do processo principal
    os.chdir(project_root)
    print(f"📂 Working dir forçado: {Path.cwd()}")

    if not process_pdf_sync:
        print("⚠️ core_xylella não disponível.")
        excel_path = stable_copy.with_suffix(".xlsx")
        return [{"path": str(excel_path), "processed": 0, "discrepancy": False}]

    print(f"🧪 A processar: {stable_copy.name}")
    result = process_pdf_sync(str(stable_copy))

    # ✅ Confirmar se debug/ e summary existem
    debug_dir = project_root / "debug"
    summary_files = list(debug_dir.glob("*_summary.txt")) if debug_dir.exists() else []
    if summary_files:
        print(f"🧾 Summary encontrado: {summary_files[-1]}")
    else:
        print("⚠️ Nenhum summary encontrado no diretório do projeto!")

    return _normalize_result(result)


def _normalize_result(result):
    """Normaliza diferentes formatos devolvidos pelo core."""
    entries = []
    if isinstance(result, list):
        for r in result:
            if isinstance(r, str):
                entries.append({"path": r, "processed": 0, "discrepancy": False})
            elif isinstance(r, dict):
                entries.append(r)
            elif isinstance(r, tuple):
                entries.append({
                    "path": r[0],
                    "processed": r[1] if len(r) > 1 else 0,
                    "discrepancy": bool(r[2]) if len(r) > 2 else False
                })
    elif isinstance(result, tuple):
        files, samples, discrepancies = result
        for i, f in enumerate(files):
            entries.append({
                "path": str(f),
                "processed": samples if isinstance(samples, int) else samples[i] if isinstance(samples, list) else 0,
                "discrepancy": discrepancies if isinstance(discrepancies, bool) else bool(discrepancies[i]) if isinstance(discrepancies, list) else False
            })
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
