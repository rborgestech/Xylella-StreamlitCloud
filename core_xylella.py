# -*- coding: utf-8 -*-
"""
core_xylella.py — Cloud/Streamlit (OCR Azure + Parser Colab + Writer por requisição)

API esperada pela UI (xylella_processor.py):
    • process_pdf_sync(pdf_path) -> List[List[Dict]]   # devolve listas de amostras por requisição
    • write_to_template(rows_per_req, out_base_path, expected_count=None, source_pdf=None) -> List[str]

Requer:
  - AZURE_API_KEY, AZURE_ENDPOINT (env)
  - TEMPLATE_PATH (env) ou ficheiro 'TEMPLATE_PXf_SGSLABIP1056.xlsx' ao lado do core
  - OUTPUT_DIR (env) — diretório onde guardar .xlsx e _ocr_debug.txt
"""

from __future__ import annotations
import os, re, io, time, json, shutil, requests
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Any, Tuple

from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment

# ───────────────────────────────────────────────
# Caminhos / Ambiente
# ───────────────────────────────────────────────
BASE_DIR = Path(__file__).parent
OUTPUT_DIR = Path(os.environ.get("OUTPUT_DIR", BASE_DIR / "Output"))
OUTPUT_DIR.mkdir(exist_ok=True)

TEMPLATE_PATH = Path(os.environ.get("TEMPLATE_PATH", BASE_DIR / "TEMPLATE_PXf_SGSLABIP1056.xlsx"))
if not TEMPLATE_PATH.exists():
    # não falha já — o app.py normalmente garante isto; fazemos só aviso
    print(f"ℹ️ Aviso: TEMPLATE não encontrado em {TEMPLATE_PATH}. Será verificado no momento da escrita.")

AZURE_API_KEY = os.environ.get("AZURE_API_KEY", "")
AZURE_ENDPOINT = os.environ.get("AZURE_ENDPOINT", "")
MODEL_ID = os.environ.get("AZURE_MODEL_ID", "prebuilt-document")

# ───────────────────────────────────────────────
# Estilos Excel
# ───────────────────────────────────────────────
GREEN = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
RED   = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
GRAY  = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
BOLD  = Font(bold=True, color="000000")
ITALIC= Font(italic=True, color="555555")

# ───────────────────────────────────────────────
# Utilitários
# ───────────────────────────────────────────────
def _is_valid_date(v) -> bool:
    try:
        datetime.strptime(str(v).strip(), "%d/%m/%Y")
        return True
    except Exception:
        return False

def _to_dt(v):
    try:
        return datetime.strptime(str(v).strip(), "%d/%m/%Y")
    except Exception:
        return v

def extract_all_text(result_json: Dict[str, Any]) -> str:
    lines = []
    for pg in result_json.get("analyzeResult", {}).get("pages", []):
        for ln in pg.get("lines", []):
            txt = (ln.get("content") or ln.get("text") or "").strip()
            if txt:
                lines.append(txt)
    return "\n".join(lines)

# ───────────────────────────────────────────────
# OCR Azure (PDF direto)
# ───────────────────────────────────────────────
def azure_analyze_pdf(pdf_path: str) -> Dict[str, Any]:
    if not AZURE_API_KEY or not AZURE_ENDPOINT:
        raise RuntimeError("Azure não configurado (AZURE_API_KEY/AZURE_ENDPOINT).")

    url = f"{AZURE_ENDPOINT.rstrip('/')}/formrecognizer/documentModels/{MODEL_ID}:analyze?api-version=2023-07-31"
    headers = {"Ocp-Apim-Subscription-Key": AZURE_API_KEY, "Content-Type": "application/pdf"}

    with open(pdf_path, "rb") as f:
        resp = requests.post(url, data=f.read(), headers=headers, timeout=90)
    if resp.status_code != 202:
        raise RuntimeError(f"Azure analyze falhou: {resp.status_code} {resp.text}")

    op = resp.headers.get("Operation-Location")
    if not op:
        raise RuntimeError("Azure não devolveu Operation-Location.")

    for _ in range(60):
        time.sleep(1.2)
        r = requests.get(op, headers={"Ocp-Apim-Subscription-Key": AZURE_API_KEY}, timeout=30)
        j = r.json()
        if j.get("status") == "succeeded":
            return j
        if j.get("status") == "failed":
            raise RuntimeError(f"OCR Azure falhou: {j}")
    raise RuntimeError("Timeout a aguardar OCR Azure.")

# ───────────────────────────────────────────────
# Parser — Secções do Colab (adaptadas)
# ───────────────────────────────────────────────
NATUREZA_KEYWORDS = [
    "ramos","folhas","ramosefolhas","ramosc/folhas","material","materialherbalho",
    "materialherbário","materialherbalo","natureza","insetos","sementes","solo"
]
TIPO_RE = re.compile(r"\b(Simples|Composta|Individual)\b", re.I)

def _looks_like_natureza(txt: str) -> bool:
    t = re.sub(r"\s+", "", (txt or "").lower())
    return any(k in t for k in NATUREZA_KEYWORDS)

def _clean_ref(raw: str) -> str:
    s = (raw or "").strip()
    s = re.sub(r"\s*/\s*", "/", s)
    s = re.sub(r"/{2,}", "/", s)
    s = re.sub(r"[A-Za-z]+", lambda m: m.group(0).upper(), s)
    s = s.replace("LUT", "LVT")
    s = re.sub(r"\s+", "", s)
    s = re.sub(r"[^A-Z0-9/]+$", "", s)
    return s

def detect_requisicoes(full_text: str) -> Tuple[int, list[int]]:
    pat = re.compile(
        r"PROGRAMA\s+NACIONAL\s+DE\s+PROSPE[ÇC][AÃ]O\s+DE\s+PRAGAS\s+DE\s+QUARENTENA",
        re.I,
    )
    matches = list(pat.finditer(full_text))
    if not matches:
        print("🔍 Nenhum cabeçalho encontrado — assumido 1 requisição.")
        return 1, []
    positions = [m.start() for m in matches]
    print(f"🔍 Detetadas {len(matches)} requisições (posições: {positions})")
    return len(matches), positions

def split_if_multiple_requisicoes(full_text: str) -> List[str]:
    text = re.sub(r"[ \t]+", " ", full_text)
    text = re.sub(r"\n{2,}", "\n", text)
    pat = re.compile(
        r"(?:PROGRAMA\s+NACIONAL\s+DE\s+PROSPE[ÇC][AÃ]O\s+DE\s+PRAGAS\s+DE\s+QUARENTENA)",
        re.I,
    )
    marks = [m.start() for m in pat.finditer(text)]
    if not marks or len(marks) == 1:
        print("🔍 Documento tratado como 1 requisição.")
        return [text]
    marks.append(len(text))
    blocos = []
    for i in range(len(marks) - 1):
        start = max(0, marks[i] - 200)
        end = min(len(text), marks[i + 1] + 200)
        bloco = text[start:end].strip()
        if len(bloco) > 400:
            blocos.append(bloco)
    print(f"📄 Documento dividido em {len(blocos)} requisições.")
    return blocos

def extract_context_from_text(full_text: str) -> Dict[str, Any]:
    ctx: Dict[str, Any] = {}
    # Zona
    m_zona = re.search(r"Xylella\s+fastidiosa\s*\(([^)]+)\)", full_text, re.I)
    ctx["zona"] = m_zona.group(1).strip() if m_zona else "Zona Isenta"

    # DGAV / Responsável
    responsavel, dgav = None, None
    m_hdr = re.search(
        r"Amostra(?:s|\(s\))?\s*colhida(?:s|\(s\))?\s*por\s*DGAV\s*[:\-]?\s*(.*)",
        full_text, re.I,
    )
    if m_hdr:
        tail = full_text[m_hdr.end():]
        linhas = [m_hdr.group(1)] + tail.splitlines()
        for ln in linhas[:4]:
            ln = (ln or "").strip()
            if ln:
                responsavel = ln; break
        if responsavel:
            responsavel = re.sub(r"\S+@dgav\.pt|\S+@\S+", "", responsavel, flags=re.I)
            responsavel = re.sub(r"PROGRAMA.*|Data.*|N[º°].*", "", responsavel, flags=re.I)
            responsavel = re.sub(r"[:;,.\-–—]+$", "", responsavel).strip()
    if responsavel:
        dgav = responsavel if re.match(r"^DGAV\b", responsavel, re.I) else f"DGAV {responsavel}"
    else:
        m_d = re.search(r"\bDGAV(?:\s+[A-Za-zÀ-ÿ?]+){1,4}", full_text)
        dgav = re.sub(r"[:;,.\-–—]+$", "", m_d.group(0)).strip() if m_d else "DGAV"
    ctx["dgav"] = dgav
    ctx["responsavel_colheita"] = None

    # Datas de colheita (*) / (**)
    colheita_map: Dict[str, str] = {}
    for m in re.finditer(r"(\d{1,2}/\d{1,2}/\d{4})\s*\(\s*(\*+)\s*\)", full_text):
        colheita_map[f"({m.group(2).replace(' ', '')})"] = m.group(1)
    if not colheita_map:
        m_simple = re.search(r"Data\s+de\s+colheita\s*[:\-\s]*([0-9/\-\s]+)", full_text, re.I)
        if m_simple:
            only_date = re.sub(r"\s+", "", m_simple.group(1))
            for key in ("(*)", "(**)", "(***)"):
                colheita_map[key] = only_date
    ctx["colheita_map"] = colheita_map
    ctx["default_colheita"] = next(iter(colheita_map.values()), "")

    # Data de envio
    m_envio = re.search(
        r"Data\s+(?:do|de)\s+envio(?:\s+ao\s+laborat[oó]rio)?[:\-\s]*([0-9/\-\s]+)",
        full_text, re.I,
    )
    if m_envio:
        ctx["data_envio"] = re.sub(r"\s+", "", m_envio.group(1))
    elif ctx["default_colheita"]:
        ctx["data_envio"] = ctx["default_colheita"]
    else:
        ctx["data_envio"] = datetime.now().strftime("%d/%m/%Y")

    # Nº amostras declaradas (se houver no formulário)
    flat = re.sub(r"\s+", " ", full_text)
    m_decl = re.search(r"N[º°]?\s*de\s*amostras(?:\s+neste\s+envio)?\s*[:\-]?\s*(\d{1,4})", flat, re.I)
    ctx["declared_samples"] = int(m_decl.group(1)) if m_decl else None

    return ctx

def parse_xylella_tables(result_json: Dict[str, Any], context: Dict[str, Any], req_id=None) -> List[Dict[str, Any]]:
    """Extrai amostras das tabelas Azure, aplicando o contexto de cada requisição."""
    out: List[Dict[str, Any]] = []
    tables = result_json.get("analyzeResult", {}).get("tables", []) or []
    if not tables:
        print("⚠️ Nenhuma tabela encontrada.")
        return out

    def clean_value(x: Any) -> str:
        if x is None: return ""
        s = str(x).replace("\n", " ").strip()
        return re.sub(r"\s{2,}", " ", s)

    for t in tables:
        cells = t.get("cells", [])
        if not cells: 
            continue
        nc = max(c.get("columnIndex", 0) for c in cells) + 1
        nr = max(c.get("rowIndex", 0) for c in cells) + 1
        grid = [[""] * nc for _ in range(nr)]
        for c in cells:
            r, ci = c.get("rowIndex", 0), c.get("columnIndex", 0)
            if r < 0 or ci < 0: 
                continue
            if r >= len(grid):
                grid.extend([[""] * nc for _ in range(r - len(grid) + 1)])
            grid[r][ci] = clean_value(c.get("content", ""))

        for row in grid:
            if not row or not any(row):
                continue
            ref_raw = (row[0] if len(row) > 0 else "").strip()
            ref = _clean_ref(ref_raw)
            if not ref or re.match(r"^\D+$", ref):
                continue

            hospedeiro = row[2] if len(row) > 2 else ""
            obs        = row[3] if len(row) > 3 else ""

            if _looks_like_natureza(hospedeiro):
                hospedeiro = ""

            tipo = ""
            joined = " ".join([x for x in row if isinstance(x, str)])
            m_tipo = TIPO_RE.search(joined)
            if m_tipo:
                tipo = m_tipo.group(1).capitalize()
                obs = TIPO_RE.sub("", obs).strip()

            datacolheita = context.get("default_colheita", "")
            m_ast = re.search(r"\(\s*\*+\s*\)", joined)
            if m_ast:
                mark = re.sub(r"\s+", "", m_ast.group(0))
                datacolheita = context.get("colheita_map", {}).get(mark, datacolheita)

            if obs.strip().lower() in ("simples", "composta", "individual"):
                obs = ""

            out.append({
                "requisicao_id": req_id,
                "datarececao": context["data_envio"],
                "datacolheita": datacolheita,
                "referencia": ref,
                "hospedeiro": hospedeiro,
                "tipo": tipo,
                "zona": context["zona"],
                "responsavelamostra": context["dgav"],
                "responsavelcolheita": context["responsavel_colheita"],
                "observacoes": obs,
                "procedure": "XYLELLA",
                "datarequerido": context["data_envio"],
                "Score": ""
            })

    print(f"✅ {len(out)} amostras extraídas (req_id={req_id}).")
    return out

def parse_all_requisitions(result_json: Dict[str, Any], pdf_name: str, txt_path: str | None) -> List[List[Dict[str, Any]]]:
    """Divide o documento em blocos (requisições) e extrai amostras por bloco."""
    # Texto global
    if txt_path and os.path.exists(txt_path):
        full_text = Path(txt_path).read_text(encoding="utf-8")
        print(f"📝 Contexto extraído de {os.path.basename(txt_path)}")
    else:
        full_text = extract_all_text(result_json)

    count, _ = detect_requisicoes(full_text)
    all_tables = result_json.get("analyzeResult", {}).get("tables", []) or []

    # Documento simples
    if count <= 1:
        context = extract_context_from_text(full_text)
        amostras = parse_xylella_tables(result_json, context, req_id=1)
        return [amostras] if amostras else []

    # Múltiplas requisições
    blocos = split_if_multiple_requisicoes(full_text)
    out: List[List[Dict[str, Any]]] = []

    for i, bloco in enumerate(blocos, start=1):
        try:
            context = extract_context_from_text(bloco)

            # refs usadas para filtrar tabelas do Azure para este bloco
            refs_bloco = re.findall(r"\b\d{7,8}\b|\b\d{2,4}/\d{2,4}/[A-Z0-9\-]+\b", bloco, re.I)
            tables_filtradas = []
            for t in all_tables:
                joined = " ".join(c.get("content", "") for c in t.get("cells", []))
                if any(ref in joined for ref in refs_bloco):
                    tables_filtradas.append(t)

            if not tables_filtradas:
                print(f"⚠️ Sem correspondência de tabelas na requisição {i}. Ignorado.")
                continue

            local = {"analyzeResult": {"tables": tables_filtradas}}
            amostras = parse_xylella_tables(local, context, req_id=i)
            if amostras:
                out.append(amostras)
        except Exception as e:
            print(f"❌ Erro na requisição {i}: {e}")

    return out

# ───────────────────────────────────────────────
# Escrita no TEMPLATE — 1 ficheiro por requisição
# ───────────────────────────────────────────────
def write_to_template(rows_per_req: List[List[Dict[str, Any]]],
                      out_base_path: str,
                      expected_count: int | None = None,
                      source_pdf: str | None = None) -> List[str]:
    """Gera vários ficheiros: <out_base>_req1.xlsx, _req2.xlsx, ..."""
    template_path = Path(os.environ.get("TEMPLATE_PATH", TEMPLATE_PATH))
    if not template_path.exists():
        raise FileNotFoundError(f"TEMPLATE não encontrado: {template_path}")
    out_dir = Path(os.environ.get("OUTPUT_DIR", OUTPUT_DIR))
    out_dir.mkdir(exist_ok=True)

    sheet = "Avaliação pré registo"
    start_row = 6
    base = Path(out_base_path).stem
    out_files: List[str] = []

    effective = rows_per_req if rows_per_req else [[]]

    for idx, req in enumerate(effective, start=1):
        # copia física do template para isolar o workbook (sem cache de openpyxl)
        tmp_template = out_dir / f"__tmp_req{idx}.xlsx"
        shutil.copy(template_path, tmp_template)
        wb = load_workbook(tmp_template)
        if sheet not in wb.sheetnames:
            wb.close(); tmp_template.unlink(missing_ok=True)
            raise KeyError(f"Folha '{sheet}' não encontrada no template.")
        ws = wb[sheet]

        # 🧹 limpar área de dados (sem tocar cabeçalhos/fórmulas)
        for r in range(start_row, ws.max_row + 1):
            for c in range(1, 13):
                ws.cell(row=r, column=c).value = None

        # ✍️ escrever linhas
        for ridx, row in enumerate(req, start=start_row):
            A, B = ws[f"A{ridx}"], ws[f"B{ridx}"]
            rece, colh = row.get("datarececao", ""), row.get("datacolheita", "")
            A.value = _to_dt(rece) if _is_valid_date(rece) else rece
            B.value = _to_dt(colh) if _is_valid_date(colh) else colh
            if not _is_valid_date(rece): A.fill = RED
            if not _is_valid_date(colh): B.fill = RED

            ws[f"C{ridx}"] = row.get("referencia", "")
            ws[f"D{ridx}"] = row.get("hospedeiro", "")
            ws[f"E{ridx}"] = row.get("tipo", "")
            ws[f"F{ridx}"] = row.get("zona", "")
            ws[f"G{ridx}"] = row.get("responsavelamostra", "")
            ws[f"H{ridx}"] = row.get("responsavelcolheita", "")
            ws[f"I{ridx}"] = ""  # Observações
            ws[f"K{ridx}"] = row.get("procedure", "XYLELLA")
            ws[f"L{ridx}"] = f"=A{ridx}+30"

            # obrigatórios A→G: cor se vazio
            for col in ("A","B","C","D","E","F","G"):
                cell = ws[f"{col}{ridx}"]
                if cell.value is None or str(cell.value).strip() == "":
                    cell.fill = RED

        processed = len(req)

        # E1:F1 — total processado (e, se quiseres, podes comparar com expected_count)
        ws.merge_cells("E1:F1")
        ws["E1"].value = f"Nº Amostras: {processed}"
        ws["E1"].font = BOLD
        ws["E1"].alignment = Alignment(horizontal="center", vertical="center")
        ws["E1"].fill = GREEN if processed > 0 else RED

        # G1:J1 — origem do PDF
        ws.merge_cells("G1:J1")
        ws["G1"].value = f"Origem: {os.path.basename(source_pdf) if source_pdf else base}"
        ws["G1"].font = ITALIC
        ws["G1"].alignment = Alignment(horizontal="left", vertical="center")
        ws["G1"].fill = GRAY

        # K1:L1 — timestamp
        ws.merge_cells("K1:L1")
        ws["K1"].value = f"Processado em: {datetime.now():%d/%m/%Y %H:%M}"
        ws["K1"].font = ITALIC
        ws["K1"].alignment = Alignment(horizontal="right", vertical="center")
        ws["K1"].fill = GRAY

        out_path = out_dir / f"{base}_req{idx}.xlsx"
        wb.save(out_path); wb.close()
        tmp_template.unlink(missing_ok=True)
        print(f"🟢 Gravado: {out_path}")
        out_files.append(str(out_path))

    return out_files

# ───────────────────────────────────────────────
# OCR + Parsing (devolve listas por requisição)
# ───────────────────────────────────────────────
def process_pdf_sync(pdf_path: str) -> List[List[Dict[str, Any]]]:
    pdf_path = str(pdf_path)
    base = Path(pdf_path).stem

    # OCR Azure
    result_json = azure_analyze_pdf(pdf_path)

    # Guardar texto OCR bruto para debug
    full_text = extract_all_text(result_json)
    debug_txt = OUTPUT_DIR / f"{base}_ocr_debug.txt"
    with open(debug_txt, "w", encoding="utf-8") as f:
        f.write(full_text)
    print(f"📝 Texto OCR bruto guardado em: {debug_txt}")

    # Parser completo (Colab): tabelas + contexto + split
    rows_per_req = parse_all_requisitions(result_json, pdf_path, str(debug_txt))
    total = sum(len(r) for r in rows_per_req)
    print(f"🏁 Concluído: {len(rows_per_req)} requisições, {total} amostras.")
    return rows_per_req

# Execução direta (debug local)
if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="Xylella Processor (Azure OCR + Parser Colab + Writer)")
    ap.add_argument("pdf", help="PDF a processar")
    ap.add_argument("--expected", type=int, default=None)
    args = ap.parse_args()

    rows = process_pdf_sync(args.pdf)
    base = Path(args.pdf).stem
    write_to_template(rows, base, expected_count=args.expected, source_pdf=Path(args.pdf).name)
