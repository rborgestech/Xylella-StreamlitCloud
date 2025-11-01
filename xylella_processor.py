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
    - Cria pasta debug/ e ficheiro summary.
    - Retorna lista de dicionários: [{path, processed, discrepancy}, ...]
    """
    pdf_path = Path(pdf_path).resolve()
    pdf_name = pdf_path.stem
    debug_dir = Path.cwd() / "debug"
    debug_dir.mkdir(exist_ok=True)

    if not process_pdf_sync:
        print("⚠️ core_xylella não encontrado — devolve simulação.")
        excel_path = pdf_path.with_suffix(".xlsx")
        return [{"path": str(excel_path), "processed": 0, "discrepancy": False}]

    print(f"🧪 Início de processamento: {pdf_path.name}")
    result = process_pdf_sync(str(pdf_path))

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

    # Cria summary.txt dentro de debug
    summary_path = debug_dir / f"{pdf_name}_summary.txt"
    with open(summary_path, "w", encoding="utf-8") as f:
        f.write(f"🧾 RESUMO DE EXECUÇÃO — {datetime.now():%Y-%m-%d %H:%M:%S}\n")
        f.write(f"PDF: {pdf_path.name}\n\n")

        total_amostras = 0
        discrep = 0
        for e in entries:
            base = Path(e["path"]).name
            proc = e.get("processed") or 0
            disc = e.get("discrepancy")
            if disc:
                discrep += 1
                f.write(f"⚠️ {base}: ficheiro gerado com discrepância.\n")
            else:
                f.write(f"✅ {base}: ficheiro gerado. ({proc} amostras OK)\n")
            total_amostras += proc

        f.write(f"\n📊 Total de ficheiros: {len(entries)}\n")
        f.write(f"🧪 Total de amostras processadas: {total_amostras}\n")
        f.write(f"⚠️ Ficheiros com discrepâncias: {discrep}\n")

    print(f"✅ Ficheiro summary criado em {summary_path}")
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
