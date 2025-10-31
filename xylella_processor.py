# xylella_processor.py
import os, io, zipfile
from pathlib import Path
from typing import List, Tuple, Dict
import importlib

_CORE_MODULE_NAME = "core_xylella"
core = importlib.import_module(_CORE_MODULE_NAME)

OUTPUT_DIR = Path(os.environ.get("OUTPUT_DIR", Path(__file__).parent / "Output"))
OUTPUT_DIR.mkdir(exist_ok=True)

def process_pdf(pdf_path: str) -> List[str]:
    """Mantido por compatibilidade — só devolve caminhos dos .xlsx."""
    files, _ = process_pdf_with_stats(pdf_path)
    return files

def process_pdf_with_stats(pdf_path: str) -> Tuple[List[str], Dict]:
    """
    Processa um PDF e devolve:
      - lista de paths dos .xlsx gerados (um por requisição)
      - dicionário 'stats' com contagens fiáveis (por req e total)
    """
    rows_per_req = core.process_pdf_sync(pdf_path)  # List[List[Dict]]
    base_name = os.path.splitext(os.path.basename(pdf_path))[0]

    created: List[str] = []
    per_req = []
    single = len(rows_per_req) == 1

    for i, rows in enumerate(rows_per_req, start=1):
        # 1 req → nome simples; >1 req → _req{i}
        out_name = f"{base_name}.xlsx" if single else f"{base_name}_req{i}.xlsx"
        out_path = core.write_to_template(
            rows,
            out_name,
            expected_count=None,           # podes ligar ao context["declared_samples"] se quiseres
            source_pdf=pdf_path
        )
        if out_path:
            created.append(out_path)
            per_req.append({
                "req": i,
                "samples": len(rows),      # ✅ nº real de amostras nesta requisição
                "expected": None,          # (opcional) preenche se extraíres do contexto
                "file": out_path
            })

    stats = {
        "pdf": os.path.basename(pdf_path),
        "req_count": len(rows_per_req),
        "samples_total": sum(len(r) for r in rows_per_req),
        "per_req": per_req
    }
    return created, stats

def write_to_template(rows, out_base_path, expected_count=None, source_pdf=None):
    return core.write_to_template(rows, out_base_path, expected_count, source_pdf)

def build_zip(file_paths: List[str]) -> bytes:
    """ZIP simples (sem summary)."""
    mem = io.BytesIO()
    with zipfile.ZipFile(mem, "w", compression=zipfile.ZIP_DEFLATED) as z:
        for p in file_paths:
            if p and os.path.exists(p):
                z.write(p, arcname=os.path.basename(p))
    mem.seek(0)
    return mem.read()

def build_zip_with_summary(file_paths: List[str], summary_data: List[Dict]) -> bytes:
    """
    Cria um ZIP com todos os ficheiros e summary.txt (com contagens corretas).
    """
    mem = io.BytesIO()
    with zipfile.ZipFile(mem, "w", compression=zipfile.ZIP_DEFLATED) as z:
        lines = []
        for s in summary_data:
            lines.append(f"📄 {s['pdf']}: {s['req_count']} requisições, {s['samples_total']} amostras.")
            for req in s.get("per_req", []):
                line = f"  • Requisição {req['req']}: {req['samples']} amostras"
                if req.get("expected") is not None:
                    diff = req['samples'] - (req['expected'] or 0)
                    sign = "+" if diff > 0 else ""
                    line += f" / {req['expected']} esperadas"
                    if diff != 0:
                        line += f" ⚠️ ({sign}{diff} diferença)"
                line += f" → {Path(req['file']).name}"
                lines.append(line)
            lines.append("")
        z.writestr("summary.txt", "\n".join(lines))

        for f in file_paths:
            if f and os.path.exists(f):
                z.write(f, arcname=Path(f).name)

    mem.seek(0)
    return mem.read()
