# xylella_processor.py
import os, io, zipfile
from pathlib import Path
from typing import List, Tuple, Dict
import importlib

_CORE_MODULE_NAME = "core_xylella"
core = importlib.import_module(_CORE_MODULE_NAME)

OUTPUT_DIR = Path(os.environ.get("OUTPUT_DIR", Path(__file__).parent / "Output"))
OUTPUT_DIR.mkdir(exist_ok=True)

# ───────────────────────────────────────────────
# PROCESSAMENTO COM ESTATÍSTICAS
# ───────────────────────────────────────────────
def process_pdf_with_stats(pdf_path: str) -> Tuple[List[str], Dict]]:
    """
    Processa um PDF e devolve:
      - lista de paths dos .xlsx gerados (um por requisição)
      - dicionário 'stats' com contagens fiáveis (por req e total)
    Suporta contexto com número de amostras declaradas (expected_count).
    """
    # O core devolve agora [{ "rows": [...], "expected": int|None }, ...]
    rows_per_req = core.process_pdf_sync(pdf_path)

    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
    created: List[str] = []
    per_req = []
    single = len(rows_per_req) == 1

    for i, req in enumerate(rows_per_req, start=1):
        if isinstance(req, dict):
            rows = req.get("rows", [])
            expected = req.get("expected")
        else:
            rows = req
            expected = None

        # Nome do ficheiro
        out_name = f"{base_name}.xlsx" if single else f"{base_name}_req{i}.xlsx"

        # Grava template com validação esperadas/processadas
        out_path = core.write_to_template(
            rows,
            out_name,
            expected_count=expected,
            source_pdf=pdf_path
        )

        if out_path:
            created.append(out_path)
            per_req.append({
                "req": i,
                "samples": len(rows),
                "expected": expected,
                "file": out_path
            })

    stats = {
        "pdf": os.path.basename(pdf_path),
        "req_count": len(rows_per_req),
        "samples_total": sum(len(r.get("rows", r)) if isinstance(r, dict) else len(r) for r in rows_per_req),
        "per_req": per_req
    }

    return created, stats

# ───────────────────────────────────────────────
# COMPATIBILIDADE ANTIGA
# ───────────────────────────────────────────────
def process_pdf(pdf_path: str) -> List[str]:
    files, _ = process_pdf_with_stats(pdf_path)
    return files

def write_to_template(rows, out_base_path, expected_count=None, source_pdf=None):
    return core.write_to_template(rows, out_base_path, expected_count, source_pdf)

# ───────────────────────────────────────────────
# ZIP COM SUMMARY.TXT
# ───────────────────────────────────────────────
def build_zip_with_summary(file_paths: List[str], summary_data: List[Dict]) -> bytes:
    """
    Cria um ZIP com todos os ficheiros Excel e um summary.txt detalhado,
    incluindo diferenças esperadas/processadas por requisição.
    """
    mem = io.BytesIO()
    with zipfile.ZipFile(mem, "w", compression=zipfile.ZIP_DEFLATED) as z:
        lines = []
        for s in summary_data:
            lines.append(f"📄 {s['pdf']}: {s['req_count']} requisições, {s['samples_total']} amostras.")
            for req in s.get("per_req", []):
                line = f"  • Requisição {req['req']}: {req['samples']} amostras"
                expected = req.get("expected")
                if expected is not None:
                    diff = req["samples"] - expected
                    sign = "+" if diff > 0 else ""
                    line += f" / {expected} esperadas"
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

# ───────────────────────────────────────────────
# ZIP SIMPLES (SEM SUMMARY)
# ───────────────────────────────────────────────
def build_zip(file_paths: List[str]) -> bytes:
    mem = io.BytesIO()
    with zipfile.ZipFile(mem, "w", compression=zipfile.ZIP_DEFLATED) as z:
        for p in file_paths:
            if p and os.path.exists(p):
                z.write(p, arcname=os.path.basename(p))
    mem.seek(0)
    return mem.read()
