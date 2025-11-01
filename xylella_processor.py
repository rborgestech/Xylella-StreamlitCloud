# -*- coding: utf-8 -*-
import os
from pathlib import Path
from datetime import datetime

# Import seguro do core real
try:
    from core_xylella import process_pdf_sync
except ImportError:
    process_pdf_sync = None

def process_pdf(pdf_path):
    """
    Wrapper que invoca o processador real (core_xylella).
    Cria automaticamente pasta debug/ e summary.
    """
    if not process_pdf_sync:
        print("⚠️ core_xylella não encontrado — devolve lista simulada.")
        excel_path = Path(pdf_path).with_suffix(".xlsx")
        return [{"path": str(excel_path), "processed": 0, "discrepancy": False}]

    pdf_name = Path(pdf_path).stem
    debug_dir = Path.cwd() / "debug"
    debug_dir.mkdir(exist_ok=True)

    # Executa o core
    print(f"🧪 Início de processamento: {Path(pdf_path).name}")
    result = process_pdf_sync(pdf_path)

    # result pode ser lista de paths ou lista de dicts
    entries = []
    if isinstance(result, list):
        for r in result:
            if isinstance(r, str):
                entries.append({"path": r, "processed": None, "discrepancy": False})
            elif isinstance(r, dict):
                entries.append(r)
            elif isinstance(r, tuple) and len(r) >= 1:
                entries.append({"path": r[0], "processed": None, "discrepancy": False})

    # Cria summary.txt
    summary_path = debug_dir / f"{pdf_name}_summary.txt"
    with open(summary_path, "w", encoding="utf-8") as f:
        f.write(f"🧾 RESUMO DE EXECUÇÃO — {datetime.now():%Y-%m-%d %H:%M:%S}\n")
        f.write(f"PDF: {Path(pdf_path).name}\n\n")
        total_amostras = 0
        discrep_count = 0
        for e in entries:
            base = Path(e['path']).name
            proc = e.get("processed") or 0
            discrep = e.get("discrepancy")
            if discrep:
                discrep_count += 1
                f.write(f"⚠️ {base}: {proc} amostras (discrepância)\n")
            else:
                f.write(f"✅ {base}: {proc} amostras OK\n")
            total_amostras += proc
        f.write(f"\n📊 Total de ficheiros: {len(entries)}\n")
        f.write(f"🧪 Total de amostras processadas: {total_amostras}\n")
        f.write(f"⚠️ Ficheiros com discrepâncias: {discrep_count}\n")

    print(f"✅ Ficheiro summary criado em {summary_path}")
    return entries


def build_zip(paths):
    """
    Gera um ZIP com os paths fornecidos.
    """
    import io, zipfile
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as z:
        for p in paths:
            p = Path(p)
            if p.exists():
                z.write(p, arcname=p.name)
    zip_buffer.seek(0)
    return zip_buffer.getvalue()
