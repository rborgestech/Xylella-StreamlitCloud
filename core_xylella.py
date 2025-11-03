# -*- coding: utf-8 -*-
"""
core_xylella.py — versão estável com cache + paralelismo
Mantém todas as funcionalidades originais (parser, Excel, OCR) e adiciona:
 - Cache local do OCR Azure (por hash do ficheiro)
 - Processamento paralelo das requisições
"""

import os
import re
import time
import json
import hashlib
import tempfile
import requests
from datetime import datetime
from pathlib import Path
from typing import Dict, Any, List
from concurrent.futures import ThreadPoolExecutor, as_completed
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment

# ───────────────────────────────────────────────
# Diretórios e variáveis
# ───────────────────────────────────────────────
OUTPUT_DIR = Path(os.getenv("OUTPUT_DIR", "output_final"))
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
CACHE_DIR = OUTPUT_DIR / "ocr_cache"
CACHE_DIR.mkdir(parents=True, exist_ok=True)
MAX_REQ_WORKERS = 4

TEMPLATE_PATH = Path(os.environ.get("TEMPLATE_PATH", Path(__file__).parent / "TEMPLATE_PXf_SGSLABIP1056.xlsx"))

AZURE_API_KEY = os.getenv("AZURE_API_KEY", "")
AZURE_ENDPOINT = os.getenv("AZURE_ENDPOINT", "")
MODEL_ID = os.getenv("AZURE_MODEL_ID", "prebuilt-document")

# ───────────────────────────────────────────────
# Cache OCR
# ───────────────────────────────────────────────
def _hash_file(pdf_path: str) -> str:
    h = hashlib.sha1()
    with open(pdf_path, "rb") as f:
        while chunk := f.read(1024 * 1024):
            h.update(chunk)
    return h.hexdigest()

def _cache_path(pdf_path: str) -> Path:
    return CACHE_DIR / f"{os.path.basename(pdf_path)}.{_hash_file(pdf_path)}.json"

def load_cached_ocr(pdf_path: str) -> Dict[str, Any] | None:
    cp = _cache_path(pdf_path)
    if cp.exists():
        try:
            print(f"♻️ OCR reutilizado: {cp.name}")
            return json.loads(cp.read_text(encoding="utf-8"))
        except Exception:
            pass
    return None

def save_cached_ocr(pdf_path: str, data: Dict[str, Any]) -> None:
    cp = _cache_path(pdf_path)
    try:
        cp.write_text(json.dumps(data), encoding="utf-8")
    except Exception:
        pass

# ───────────────────────────────────────────────
# OCR Azure
# ───────────────────────────────────────────────
def azure_analyze_pdf(pdf_path: str) -> Dict[str, Any]:
    if not AZURE_API_KEY or not AZURE_ENDPOINT:
        raise RuntimeError("Azure não configurado (AZURE_API_KEY/AZURE_ENDPOINT).")

    url = f"{AZURE_ENDPOINT.rstrip('/')}/formrecognizer/documentModels/{MODEL_ID}:analyze?api-version=2023-07-31"
    headers = {"Ocp-Apim-Subscription-Key": AZURE_API_KEY, "Content-Type": "application/pdf"}

    with open(pdf_path, "rb") as f:
        resp = requests.post(url, data=f.read(), headers=headers, timeout=120)
    if resp.status_code != 202:
        raise RuntimeError(f"Azure analyze falhou: {resp.status_code} {resp.text}")

    op = resp.headers.get("Operation-Location")
    if not op:
        raise RuntimeError("Azure não devolveu Operation-Location.")

    start = time.time()
    while True:
        r = requests.get(op, headers={"Ocp-Apim-Subscription-Key": AZURE_API_KEY}, timeout=60)
        j = r.json()
        if j.get("status") == "succeeded":
            return j
        if j.get("status") == "failed":
            raise RuntimeError(f"OCR Azure falhou: {j}")
        if time.time() - start > 180:
            raise RuntimeError("Timeout a aguardar OCR Azure.")
        time.sleep(1.2)

# ───────────────────────────────────────────────
# Helpers
# ───────────────────────────────────────────────
def extract_all_text(result_json: Dict[str, Any]) -> str:
    lines = []
    for pg in result_json.get("analyzeResult", {}).get("pages", []):
        for ln in pg.get("lines", []):
            txt = (ln.get("content") or ln.get("text") or "").strip()
            if txt:
                lines.append(txt)
    return "\n".join(lines)

# ───────────────────────────────────────────────
# Importa parser e writer originais
# ───────────────────────────────────────────────
from core_xylella_parser import (
    detect_requisicoes,
    split_if_multiple_requisicoes,
    extract_context_from_text,
    parse_xylella_tables,
    write_to_template
)

# ───────────────────────────────────────────────
# Função de parsing completa
# ───────────────────────────────────────────────
def parse_all_requisitions(result_json: Dict[str, Any], pdf_name: str, txt_path: str | None) -> List[Dict[str, Any]]:
    if txt_path and os.path.exists(txt_path):
        full_text = Path(txt_path).read_text(encoding="utf-8")
    else:
        full_text = extract_all_text(result_json)

    count, _ = detect_requisicoes(full_text)
    tables = result_json.get("analyzeResult", {}).get("tables", []) or []

    if count <= 1:
        ctx = extract_context_from_text(full_text)
        amostras = parse_xylella_tables(result_json, ctx, req_id=1)
        return [{"rows": amostras, "expected": ctx.get("declared_samples", 0)}]

    blocos = split_if_multiple_requisicoes(full_text)
    results = []
    for i, bloco in enumerate(blocos, start=1):
        ctx = extract_context_from_text(bloco)
        amostras = parse_xylella_tables({"analyzeResult": {"tables": tables}}, ctx, req_id=i)
        results.append({"rows": amostras, "expected": ctx.get("declared_samples", 0)})
    return results

# ───────────────────────────────────────────────
# Função principal
# ───────────────────────────────────────────────
def process_pdf_sync(pdf_path: str) -> List[str]:
    base = os.path.basename(pdf_path)
    print(f"\n🧪 Início de processamento: {base}")
    txt_path = OUTPUT_DIR / f"{os.path.splitext(base)[0]}_ocr_debug.txt"

    # 1️⃣ OCR com cache
    result_json = load_cached_ocr(pdf_path)
    if result_json is None:
        result_json = azure_analyze_pdf(pdf_path)
        save_cached_ocr(pdf_path, result_json)

    # 2️⃣ Guardar texto OCR
    txt_path.write_text(extract_all_text(result_json), encoding="utf-8")

    # 3️⃣ Parsing e escrita
    requisitions = parse_all_requisitions(result_json, pdf_path, str(txt_path))
    if not requisitions:
        print("⚠️ Nenhuma requisição encontrada.")
        return []

    created_files = []
    with ThreadPoolExecutor(max_workers=min(MAX_REQ_WORKERS, len(requisitions))) as executor:
        futures = {
            executor.submit(
                write_to_template,
                req["rows"],
                f"{os.path.splitext(base)[0]}_req{i}.xlsx" if len(requisitions) > 1 else f"{os.path.splitext(base)[0]}.xlsx",
                req.get("expected"),
                pdf_path,
            ): i
            for i, req in enumerate(requisitions, start=1)
        }
        for fut in as_completed(futures):
            try:
                path = fut.result()
                if path:
                    created_files.append(path)
            except Exception as e:
                print(f"❌ Erro ao gravar: {e}")

    print(f"🏁 {base}: {len(created_files)} ficheiro(s) Excel gerado(s).")
    return created_files
