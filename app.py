# -*- coding: utf-8 -*-
import os
import shutil
from pathlib import Path
from datetime import datetime

try:
    from core_xylella import process_pdf_sync
except ImportError:
    process_pdf_sync = None

# Garante que OUTPUT_DIR está sempre definido
OUTPUT_DIR = Path(os.getenv("OUTPUT_DIR", "/tmp")).resolve()
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

def process_pdf(pdf_path):
    """
    Executa o core_xylella.py no contexto real do Streamlit Cloud (/mount/src/xylella-streamlitcloud),
    e devolve uma lista de ficheiros .xlsx gerados (paths absolutos).
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

    # Criar script temporário que chama o core
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
    except Exception as e:
        print("❌ Erro ao interpretar resposta do core:", e)
        parsed = []

    # Garantir lista de strings (paths)
    if isinstance(parsed, list):
        parsed = [str(p) for p in parsed if isinstance(p, str)]
    else:
        parsed = []

    print("📁 Ficheiros gerados:", parsed)
    return parsed


def build_zip(excel_files):
    """
    Gera ZIP com os ficheiros Excel. Compatível com app.py.
    """
    import io, zipfile
    mem = io.BytesIO()
    with zipfile.ZipFile(mem, "w", zipfile.ZIP_DEFLATED) as z:
        for f in excel_files:
            if os.path.exists(f):
                z.write(f, arcname=os.path.basename(f))
        # Opcional: incluir ficheiros de debug
        for dbg in OUTPUT_DIR.glob("*_ocr_debug.txt"):
            z.write(dbg, arcname=f"debug/{dbg.name}")
        z.writestr("summary.txt", "Ficheiros processados com sucesso.")
    mem.seek(0)
    zip_name = f"xylella_output_{datetime.now():%Y%m%d_%H%M%S}.zip"
    return mem.read(), zip_name
