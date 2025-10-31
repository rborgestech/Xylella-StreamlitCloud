# xylella_processor.py — versão compatível e funcional (2025-10 Cloud)
import os, io, zipfile, re
from pathlib import Path
from typing import List, Dict, Any, Tuple
import importlib
from openpyxl import load_workbook

# ───────────────────────────────────────────────
# Importação do core_xylella
# ───────────────────────────────────────────────
core = importlib.import_module("core_xylella")

OUTPUT_DIR = Path(os.environ.get("OUTPUT_DIR", Path(__file__).parent / "Output"))
OUTPUT_DIR.mkdir(exist_ok=True)


# ───────────────────────────────────────────────
# Utilitários
# ───────────────────────────────────────────────
def _read_e1_counts(xlsx_path: str) -> Tuple[int | None, int | None]:
    """Lê a célula E1 e devolve (expected, processed)."""
    try:
        wb = load_workbook(xlsx_path, data_only=True)
        ws = wb.worksheets[0]
        val = str(ws["E1"].value or "")
        m = re.search(r"(\d+)\s*/\s*(\d+)", val)
        if m:
            return int(m.group(1)), int(m.group(2))
    except Exception:
        pass
    return None, None


def _collect_debug_files(outdir: Path) -> List[str]:
    """Recolhe ficheiros OCR e logs do core."""
    debug_files = []
    for pattern in ["*_ocr_debug.txt", "*.csv", "process_summary_*.txt", "process_log.csv"]:
        for f in outdir.glob(pattern):
            debug_files.append(str(f))
    return debug_files


# ───────────────────────────────────────────────
# Função principal
# ───────────────────────────────────────────────
def process_pdf_with_stats(pdf_path: str):
    """
    Processa o PDF com o core_xylella e devolve:
      - created_files: ficheiros Excel criados
      - stats: info por requisição
      - debug_files: logs/OCRs
    """
    print(f"📄 A processar {os.path.basename(pdf_path)} ...")
    rows_per_req = core.process_pdf_sync(pdf_path)

    # Verificação de formato — garantir listas de dicts
    if not isinstance(rows_per_req, list):
        print("⚠️ O core devolveu formato inesperado, ignorado.")
        return [], {"req_count": 0, "samples_total": 0, "per_req": []}, []

    base = os.path.splitext(os.path.basename(pdf_path))[0]
    outdir = Path(os.environ.get("OUTPUT_DIR", OUTPUT_DIR))
    created, per_req = [], []

    for i, req_data in enumerate(rows_per_req, start=1):
        # Se for texto ou vazio, ignora
        if not req_data or not isinstance(req_data, list):
            print(f"⚠️ Requisição {i} vazia ou inválida.")
            continue

        # Filtrar apenas dicionários válidos
        valid_rows = [r for r in req_data if isinstance(r, dict) and r.get("referencia")]
        if not valid_rows:
            print(f"⚠️ Requisição {i} sem amostras válidas.")
            continue

        # Gerar nome do ficheiro
        fname = f"{base}.xlsx" if len(rows_per_req) == 1 else f"{base}_req{i}.xlsx"

        # Contagem esperada
        declared = valid_rows[0].get("declared_samples") if "declared_samples" in valid_rows[0] else None

        # Escrever o Excel
        out_path = core.write_to_template(valid_rows, fname, expected_count=declared, source_pdf=pdf_path)
        if not out_path:
            continue

        expected, processed = _read_e1_counts(out_path)
        processed = processed or len(valid_rows)
        expected = expected or declared
        diff = (processed - expected) if expected is not None else None

        per_req.append({
            "req": i,
            "file": out_path,
            "samples": processed,
            "expected": expected,
            "diff": diff,
        })
        created.append(out_path)
        print(f"✅ Requisição {i}: {processed} amostras → {fname}")

    # Estatísticas
    stats = {
        "pdf_name": base,
        "req_count": len(per_req),
        "samples_total": sum(p["samples"] for p in per_req),
        "per_req": per_req,
    }

    debug_files = _collect_debug_files(outdir)
    return created, stats, debug_files


# ───────────────────────────────────────────────
# ZIP com summary e debug
# ───────────────────────────────────────────────
def build_zip_with_summary(excel_files: List[str], debug_files: List[str], summary_text: str):
    """Cria ZIP com ficheiros Excel, pasta /debug e summary.txt."""
    mem = io.BytesIO()
    zip_name = f"xylella_output_{Path.cwd().name}_{os.getpid()}.zip"

    with zipfile.ZipFile(mem, "w", zipfile.ZIP_DEFLATED) as z:
        for f in excel_files:
            if os.path.exists(f):
                z.write(f, arcname=os.path.basename(f))
        for f in debug_files:
            if os.path.exists(f):
                z.write(f, arcname=f"debug/{os.path.basename(f)}")
        z.writestr("summary.txt", summary_text or "")

    mem.seek(0)
    return mem.read(), zip_name
