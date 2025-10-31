# xylella_processor.py
# Camada fina entre a app e o core: produz XLSX, agrega estatísticas
# e recolhe os artefactos de debug para o ZIP.

from __future__ import annotations
import os, io, zipfile, re, shutil
from pathlib import Path
from typing import List, Dict, Any, Tuple
import importlib
from openpyxl import load_workbook

_CORE_MODULE_NAME = "core_xylella"
core = importlib.import_module(_CORE_MODULE_NAME)

# Onde o core vai escrever por omissão (a app pode sobrepor via env OUTPUT_DIR)
OUTPUT_DIR = Path(os.environ.get("OUTPUT_DIR", Path(__file__).parent / "Output"))
OUTPUT_DIR.mkdir(exist_ok=True)


# ─────────────────────────────────────────────────────────────────────────────
# Utilitários internos
# ─────────────────────────────────────────────────────────────────────────────
def _read_e1_counts(xlsx_path: str) -> Tuple[int|None, int|None]:
    """
    Lê a célula E1 (ex: 'Nº Amostras: 10 / 12' ou '10 / 12') e devolve (expected, processed).
    """
    try:
        wb = load_workbook(xlsx_path, data_only=True)
        ws = wb.worksheets[0]
        val = str(ws["E1"].value or "")
        # Aceita 'Nº Amostras: 10 / 12' ou '10 / 12'
        m = re.search(r"(\d+)\s*/\s*(\d+)", val)
        if m:
            expected = int(m.group(1))
            processed = int(m.group(2))
            return expected, processed
    except Exception:
        pass
    return None, None


def _ensure_dir(p: Path) -> None:
    p.mkdir(parents=True, exist_ok=True)


def _collect_debug_files(from_dir: Path) -> List[str]:
    """
    Recolhe artefactos de debug gerados pelo core no diretório OUTPUT_DIR corrente.
    """
    debug_files: List[str] = []
    patterns = ["*_ocr_debug.txt", "process_log.csv", "process_summary_*.txt"]
    for patt in patterns:
        for f in from_dir.glob(patt):
            debug_files.append(str(f))
    return debug_files


# ─────────────────────────────────────────────────────────────────────────────
# API pública
# ─────────────────────────────────────────────────────────────────────────────
def process_pdf(pdf_path: str) -> List[str]:
    """
    (Compatibilidade retro) Processa um PDF e devolve a lista de caminhos dos .xlsx gerados.
    Usa expected=None (sem validação) e nomenclatura base_reqN.xlsx.
    """
    rows_per_req = core.process_pdf_sync(pdf_path)  # List[List[Dict]]
    base = os.path.splitext(os.path.basename(pdf_path))[0]

    created: List[str] = []
    for i, rows in enumerate(rows_per_req, start=1):
        # Única requisição: sem _req1 para não sugerir que há mais.
        fname = f"{base}.xlsx" if len(rows_per_req) == 1 else f"{base}_req{i}.xlsx"
        out_path = core.write_to_template(rows, fname, expected_count=None, source_pdf=pdf_path)
        if out_path:
            created.append(out_path)
    return created


def process_pdf_with_stats(pdf_path: str) -> Tuple[List[str], Dict[str, Any], List[str]]:
    """
    Processa um PDF e devolve:
      - created_files: lista de .xlsx gerados (caminhos completos)
      - stats: { pdf_name, req_count, samples_total, per_req:[{req, file, processed, expected, diff}] }
      - debug_files: lista de artefactos de debug gerados (txt/csv)

    NOTA: Usa 'declared_samples' do contexto do core para expected_count na escrita,
    permitindo que E1 traga 'esperado/processado' correto e que a app mostre discrepâncias fiáveis.
    """
    # O core gera as listas por requisição
    rows_per_req = core.process_pdf_sync(pdf_path)  # List[List[Dict]]
    base = os.path.splitext(os.path.basename(pdf_path))[0]
    outdir = Path(os.environ.get("OUTPUT_DIR", OUTPUT_DIR))
    _ensure_dir(outdir)

    created: List[str] = []
    per_req: List[Dict[str, Any]] = []

    # Descobrir expected por requisição: o core só dá 'declared_samples' no contexto textual,
    # portanto inferimos por requisição a partir dos próprios rows (se o parser injectou essa info)
    # ou, se não houver por-linha, usamos o 'declared' do bloco (quando disponível).
    # Para não depender de metadados, primeiro gravamos e depois lemos E1 (fonte da verdade).

    for i, rows in enumerate(rows_per_req, start=1):
        fname = f"{base}.xlsx" if len(rows_per_req) == 1 else f"{base}_req{i}.xlsx"
        # tenta apanhar 'declared_samples' do bloco (se o core tiver injectado via contexto)
        declared = None
        if rows:
            # tenta em qualquer linha (o core gera todas iguais nesse campo, quando existe)
            declared = rows[0].get("declared_samples") if isinstance(rows[0], dict) else None

        out_path = core.write_to_template(rows, fname, expected_count=declared, source_pdf=pdf_path)
        if not out_path:
            continue

        created.append(out_path)
        exp, proc = _read_e1_counts(out_path)
        # Se E1 não tiver parsable, assume processed=len(rows)
        if proc is None:
            proc = len(rows)
        # Se expected não veio de E1 (porque não foi escrito), usar 'declared' (pode ser None)
        if exp is None and declared is not None:
            exp = declared

        diff = None
        if exp is not None and proc is not None:
            diff = proc - exp

        per_req.append({
            "req": i,
            "file": out_path,
            "processed": proc,
            "expected": exp,
            "diff": diff
        })

    # Totais
    samples_total = sum((r.get("processed") or 0) for r in per_req)
    stats = {
        "pdf_name": os.path.basename(pdf_path),
        "req_count": len(per_req),
        "samples_total": samples_total,
        "per_req": per_req,
    }

    # Debug gerado pelo core neste ciclo
    debug_files = _collect_debug_files(Path(outdir))

    return created, stats, debug_files


def build_zip_with_summary(
    excel_files: List[str],
    debug_files: List[str],
    summary_text: str,
    zip_prefix: str | None = None
) -> Tuple[bytes, str]:
    """
    Constrói um ZIP com:
      • ficheiros Excel na raiz
      • pasta /debug com txt/csv
      • summary.txt na raiz (conteúdo fornecido)

    Retorna (zip_bytes, zip_name).
    """
    if not zip_prefix:
        zip_prefix = "xylella_output"
    zip_name = f"{zip_prefix}.zip"

    mem = io.BytesIO()
    with zipfile.ZipFile(mem, "w", compression=zipfile.ZIP_DEFLATED) as z:
        # Excel na raiz
        for p in excel_files:
            if p and os.path.exists(p):
                z.write(p, arcname=os.path.basename(p))

        # debug/
        for p in debug_files:
            if p and os.path.exists(p):
                z.write(p, arcname=f"debug/{os.path.basename(p)}")

        # summary.txt
        z.writestr("summary.txt", summary_text or "")

    mem.seek(0)
    return mem.read(), zip_name
