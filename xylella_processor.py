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
    """Processa um PDF e devolve a lista de caminhos dos .xlsx gerados (um por requisição)."""
    rows_per_req = core.process_pdf_sync(pdf_path)
    base = os.path.splitext(os.path.basename(pdf_path))[0]
    created = []
    for i, rows in enumerate(rows_per_req, start=1):
        out_name = f"{base}_req{i}.xlsx"
        out_path = core.write_to_template(rows, out_name, expected_count=None, source_pdf=pdf_path)
        if out_path:
            created.append(out_path)
    return created

def write_to_template(rows, out_base_path, expected_count=None, source_pdf=None):
    return core.write_to_template(rows, out_base_path, expected_count, source_pdf)

def build_zip(file_paths: List[str]) -> bytes:
    """Cria um ZIP simples com os ficheiros fornecidos."""
    mem = io.BytesIO()
    with zipfile.ZipFile(mem, "w", compression=zipfile.ZIP_DEFLATED) as z:
        for p in file_paths:
            if p and os.path.exists(p):
                z.write(p, arcname=os.path.basename(p))
    mem.seek(0)
    return mem.read()

def build_zip_with_summary(file_paths: List[str], summary_data: List[dict]) -> bytes:
    """
    Cria um ZIP com todos os ficheiros e summary.txt.
    Cada entrada do summary mostra: PDF, nº de requisições, nº de amostras e discrepâncias.
    """
    mem = io.BytesIO()
    with zipfile.ZipFile(mem, "w", compression=zipfile.ZIP_DEFLATED) as z:
        summary_lines = []
        for s in summary_data:
            summary_lines.append(f"📄 {s['pdf']}: {s['req_count']} requisições, {s['samples_total']} amostras.")
            for req in s.get("per_req", []):
                line = f"  • Requisição {req['req']}: {req['samples']} amostras"
                if req.get("expected") is not None:
                    line += f" / {req['expected']} esperadas"
                    diff = req.get("samples", 0) - (req.get("expected") or 0)
                    if diff != 0:
                        sign = "+" if diff > 0 else ""
                        line += f" ⚠️ ({sign}{diff} diferença)"
                line += f" → {Path(req['file']).name}"
                summary_lines.append(line)
            summary_lines.append("")

        z.writestr("summary.txt", "\n".join(summary_lines))

        for f in file_paths:
            if f and os.path.exists(f):
                z.write(f, arcname=Path(f).name)

    mem.seek(0)
    return mem.read()
