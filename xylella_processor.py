# xylella_processor.py
import os, io, zipfile
from pathlib import Path
from typing import List
import importlib

_CORE_MODULE_NAME = "core_xylella"
core = importlib.import_module(_CORE_MODULE_NAME)

OUTPUT_DIR = Path(os.environ.get("OUTPUT_DIR", Path(__file__).parent / "Output"))
OUTPUT_DIR.mkdir(exist_ok=True)

def process_pdf(pdf_path: str) -> List[str]:
    """
    Processa um PDF e devolve a lista de caminhos dos .xlsx gerados (um por requisição).
    """
    rows_per_req = core.process_pdf_sync(pdf_path)  # List[List[Dict]]
    base = os.path.splitext(os.path.basename(pdf_path))[0]

    created = []
    for i, rows in enumerate(rows_per_req, start=1):
        out_name = f"{base}_req{i}.xlsx"
        out_path = core.write_to_template(rows, out_name, expected_count=None, source_pdf=pdf_path)
        if out_path:
            created.append(out_path)
    return created

def write_to_template(rows, out_base_path, expected_count=None, source_pdf=None):
    # apenas proxy se o app quiser usar diretamente
    return core.write_to_template(rows, out_base_path, expected_count, source_pdf)

def build_zip(file_paths: List[str]) -> bytes:
    """
    Constrói um ZIP em memória com os caminhos dados.
    """
    mem = io.BytesIO()
    with zipfile.ZipFile(mem, "w", compression=zipfile.ZIP_DEFLATED) as z:
        for p in file_paths:
            if p and os.path.exists(p):
                z.write(p, arcname=os.path.basename(p))
    mem.seek(0)
    return mem.read()
