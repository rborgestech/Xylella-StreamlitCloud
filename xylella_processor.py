# -*- coding: utf-8 -*-
import os
import shutil
from pathlib import Path
from datetime import datetime

try:
    from core_xylella import process_pdf_sync
except ImportError:
    process_pdf_sync = None

# Diretório de saída temporário (definido pelo app)
OUTPUT_DIR = Path(os.getenv("OUTPUT_DIR", "/tmp")).resolve()


def process_pdf(pdf_path):
    """
    Executa o core_xylella.py no contexto real do Streamlit Cloud (/mount/src/xylella-streamlitcloud),
    garantindo a criação de debug/ e summary. Devolve apenas a lista de paths dos ficheiros Excel.
    """
    import subprocess, json, sys

    project_root = Path("/mount/src/xylella-streamlitcloud").resolve()
    pdf_path = Path(pdf_path).resolve()
    pdf_name = pdf_path.name
    stable_pdf = project_root / pdf_name

    try:
        shutil.copy(pdf_path, stable_pdf)
    except Exception as e:
        print(f"⚠️ Erro ao copiar PDF: {e}")
        return []

    print(f"📄 Copiado para {stable_pdf}")
    print(f"📂 Working dir forçado: {project_root}")

    helper = project_root / "_run_core_wrapper.py"
    helper.write_text(f"""
import json
from core_xylella import process_pdf_sync
res = process_pdf_sync(r"{stable_pdf}")
print(json.dumps(res if isinstance(res, (list, dict)) else str(res)))
""")

    result = subprocess.run(
        [sys.executable, str(helper)],
        capture_output=True, text=True, cwd=project_root
    )

    if result.returncode != 0:
        print("❌ Erro ao executar core_xylella:")
        print(result.stderr)
        return []

    try:
        parsed = json.loads(result.stdout)
    except Exception:
        parsed = []

    entries = _normalize_result(parsed)
    return [e["path"] for e in entries]


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


def process_pdf_with_stats(pdf_path: str):
    """
    Wrapper que usa a função process_pdf e devolve stats compatíveis com o app.py.
    Garante que amostras e discrepâncias são contabilizadas corretamente.
    """
    entries = process_pdf(pdf_path)

    stats = {
        "pdf_name": os.path.basename(pdf_path),
        "req_count": len(entries),
        "samples_total": sum(e.get("processed", 0) for e in entries),
        "per_req": []
    }

    for i, e in enumerate(entries):
        stats["per_req"].append({
            "req": i + 1,
            "file": e.get("path"),
            "samples": e.get("processed", 0),
            "expected": e.get("expected"),
            "diff": e.get("processed", 0) - (e.get("expected") or 0)
        })

    # Ficheiros de debug (se existirem)
    debug_files = [str(f) for f in OUTPUT_DIR.glob("*_ocr_debug.txt")]
    return [e["path"] for e in entries], stats, debug_files



def build_zip_with_summary(excel_files, debug_files, summary_text):
    """Wrapper para manter compatibilidade com a versão do app.py que gera summary + debug."""
    import io, zipfile
    mem = io.BytesIO()
    with zipfile.ZipFile(mem, "w", zipfile.ZIP_DEFLATED) as z:
        for f in excel_files:
            if os.path.exists(f):
                z.write(f, arcname=os.path.basename(f))
        for dbg in debug_files:
            if os.path.exists(dbg):
                z.write(dbg, arcname=f"debug/{os.path.basename(dbg)}")
        z.writestr("summary.txt", summary_text or "")
    mem.seek(0)
    zip_name = f"xylella_output_{datetime.now():%Y%m%d_%H%M%S}.zip"
    return mem.read(), zip_name


# Compatibilidade com app.py
build_zip = lambda excel_files: build_zip_with_summary(excel_files, [], "Resumo do processamento gerado automaticamente.")
