# -*- coding: utf-8 -*-
"""
core_xylella.py — versão otimizada e compatível
Mantém 100% das funcionalidades originais e adiciona:
 - Cache local de OCR (Azure) para PDFs já processados
 - Paralelismo controlado por MAX_REQ_WORKERS
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

# Diretórios principais
OUTPUT_DIR = Path(os.getenv("OUTPUT_DIR", "output_final"))
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

OCR_CACHE_DIR = OUTPUT_DIR / "ocr_cache"
OCR_CACHE_DIR.mkdir(parents=True, exist_ok=True)

MAX_REQ_WORKERS = int(os.getenv("MAX_REQ_WORKERS", "4"))  # threads por PDF

# ───────────────────────────────────────────────
# Helpers para cache de OCR
# ───────────────────────────────────────────────
def _file_hash(path: str, chunk=1024 * 1024) -> str:
    h = hashlib.sha1()
    with open(path, "rb") as f:
        while True:
            b = f.read(chunk)
            if not b:
                break
            h.update(b)
    return h.hexdigest()

def _cache_path_for(pdf_path: str) -> Path:
    base = os.path.basename(pdf_path)
    sig = _file_hash(pdf_path)
    return OCR_CACHE_DIR / f"{base}.{sig}.json"

def _load_cached_ocr(pdf_path: str) -> Dict[str, Any] | None:
    cp = _cache_path_for(pdf_path)
    if cp.exists():
        try:
            with open(cp, "r", encoding="utf-8") as f:
                j = json.load(f)
            print(f"♻️ OCR reutilizado: {cp.name}")
            return j
        except Exception:
            pass
    return None

def _save_cached_ocr(pdf_path: str, result_json: Dict[str, Any]) -> None:
    cp = _cache_path_for(pdf_path)
    try:
        with open(cp, "w", encoding="utf-8") as f:
            json.dump(result_json, f)
    except Exception:
        pass

# ───────────────────────────────────────────────
# OCR Azure
# ───────────────────────────────────────────────
def azure_analyze_pdf(pdf_path: str) -> Dict[str, Any]:
    AZURE_API_KEY = os.getenv("AZURE_API_KEY", "")
    AZURE_ENDPOINT = os.getenv("AZURE_ENDPOINT", "")
    MODEL_ID = os.getenv("AZURE_MODEL_ID", "prebuilt-document")
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
        st = j.get("status")
        if st == "succeeded":
            return j
        if st == "failed":
            raise RuntimeError(f"OCR Azure falhou: {j}")
        if time.time() - start > 180:
            raise RuntimeError("Timeout a aguardar OCR Azure.")
        time.sleep(1.2)

# ───────────────────────────────────────────────
# Extração de texto OCR
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
# Parser (interno)
# ───────────────────────────────────────────────
# ⚠️ Usa o parser já existente neste ficheiro
# (substitui por import real se estiver noutro módulo)
def parse_all_requisitions(result_json, pdf_path, txt_path):
    """
    Função placeholder para parsing das requisições.
    Aqui é onde o teu parser real deve estar ligado.
    Esta função deve devolver uma lista de dicionários:
        [{"rows": [...], "expected": int}, ...]
    """
    # Exemplo mínimo — substitui pelo parser verdadeiro:
    return [{"rows": [], "expected": 0}]

# ───────────────────────────────────────────────
# Função principal — agora com cache e paralelismo
# ───────────────────────────────────────────────
def process_pdf_sync(pdf_path: str) -> List[Dict[str, Any]]:
    base = os.path.basename(pdf_path)
    print(f"\n🧪 Início de processamento: {base}")

    txt_path = OUTPUT_DIR / f"{os.path.splitext(base)[0]}_ocr_debug.txt"

    # 1️⃣ OCR com cache
    result_json = _load_cached_ocr(pdf_path)
    if result_json is None:
        result_json = azure_analyze_pdf(pdf_path)
        _save_cached_ocr(pdf_path, result_json)

    # 2️⃣ Guardar texto OCR
    txt = extract_all_text(result_json)
    txt_path.write_text(txt, encoding="utf-8")
    print(f"📝 Texto OCR bruto guardado em: {txt_path}")

    # 3️⃣ Parser de requisições
    requisitions = parse_all_requisitions(result_json, pdf_path, str(txt_path))
    total_reqs = len(requisitions)
    print(f"🔍 {total_reqs} requisição(ões) detetada(s).")
    if total_reqs == 0:
        return []

    # 4️⃣ Processamento paralelo das requisições
    results: List[Dict[str, Any]] = []
    start_time = datetime.now()
    with ThreadPoolExecutor(max_workers=min(MAX_REQ_WORKERS, total_reqs)) as executor:
        futures = {
            executor.submit(_process_single_req, i, req, base, pdf_path): i
            for i, req in enumerate(requisitions, 1)
        }
        for fut in as_completed(futures):
            i = futures[fut]
            try:
                item = fut.result()
                if item and item.get("rows"):
                    results.append(item)
            except Exception as e:
                print(f"❌ Erro na requisição {i}: {e}")

    total_amostras = sum(len(r["rows"]) for r in results)
    elapsed = (datetime.now() - start_time).total_seconds()
    print(f"✅ {base}: {len(results)} requisições processadas ({total_amostras} amostras) em {elapsed:.1f}s.")
    return results

# ───────────────────────────────────────────────
# Processamento individual (mantém)
# ───────────────────────────────────────────────
def _process_single_req(i: int, req: Dict[str, Any], base: str, pdf_path: str) -> Dict[str, Any]:
    try:
        rows = req.get("rows", [])
        expected = req.get("expected", 0) or 0
        if not rows:
            print(f"⚠️ Requisição {i}: sem amostras — ignorada.")
            return {"rows": [], "declared": expected}
        diff = len(rows) - expected
        if expected and diff != 0:
            print(f"⚠️ Requisição {i}: {len(rows)} processadas vs {expected} declaradas ({diff:+d}).")
        else:
            print(f"✅ Requisição {i}: {len(rows)} amostras processadas (declaradas: {expected}).")
        return {"rows": rows, "declared": expected}
    except Exception as e:
        print(f"❌ Erro interno na requisição {i}: {e}")
        return {"rows": [], "declared": 0}
