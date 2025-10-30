# -*- coding: utf-8 -*-
"""
core_xylella.py — Cloud/Streamlit (parser do Colab integrado)

- OCR Azure (PDF direto) com env: AZURE_API_KEY, AZURE_ENDPOINT
- Divide por requisições (cabeçalho DGAV) e extrai amostras das TABELAS Azure
- Limpa template e gera 1 Excel por requisição: <base>_req1.xlsx, ...
- Mantém E1:F1 (validação), G1:J1 (origem) e K1:L1 (timestamp)
- Respeita OUTPUT_DIR e TEMPLATE_PATH definidos no app.py
"""

from __future__ import annotations

import os, io, re, json, time
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Any, Tuple

import requests
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment

# ───────────────────────────────────────────────
# Configuração de caminhos
# ───────────────────────────────────────────────
BASE_DIR = Path(__file__).parent
OUTPUT_DIR = Path(os.environ.get("OUTPUT_DIR", BASE_DIR / "Output"))
OUTPUT_DIR.mkdir(exist_ok=True)
TEMPLATE_PATH = Path(os.environ.get("TEMPLATE_PATH", BASE_DIR / "TEMPLATE_PXf_SGSLABIP1056.xlsx"))

# Azure
AZURE_API_KEY = os.environ.get("AZURE_API_KEY", "")
AZURE_ENDPOINT = os.environ.get("AZURE_ENDPOINT", "")   # ex: https://<nome>.cognitiveservices.azure.com/
MODEL_ID = os.environ.get("AZURE_MODEL_ID", "prebuilt-document")

# Estilos Excel
GREEN  = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
RED    = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
GRAY   = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
BOLD   = Font(bold=True, color="000000")
ITALIC = Font(italic=True, color="555555")

# ───────────────────────────────────────────────
# Utilitários
# ───────────────────────────────────────────────
def _is_valid_date(v) -> bool:
    if isinstance(v, datetime): return True
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
# Azure OCR (PDF direto)
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
    for _ in range(50):
        time.sleep(1.2)
        r = requests.get(op, headers={"Ocp-Apim-Subscription-Key": AZURE_API_KEY}, timeout=30)
        j = r.json()
        if j.get("status") == "succeeded":
            return j
        if j.get("status") == "failed":
            raise RuntimeError(f"OCR Azure falhou: {j}")
    raise RuntimeError("Timeout a aguardar OCR Azure.")

# ───────────────────────────────────────────────
# Parser — (mesmo do Colab)
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
    pat = re.compile(r"PROGRAMA\s+NACIONAL\s+DE\s+PROSPE[ÇC][AÃ]O\s+DE\s+PRAGAS\s+DE\s+QUARENTENA", re.I)
    m = list(pat.finditer(full_text))
    if not m:
        print("🔍 Nenhum cabeçalho encontrado — assumido 1 requisição.")
        return 1, []
    print(f"🔍 Detetadas {len(m)} requisições (posições: {[x.start() for x in m]})")
    return len(m), [x.start() for x in m]

def split_if_multiple_requisicoes(full_text: str) -> List[str]:
    text = re.sub(r"[ \t]+", " ", full_text)
    text = re.sub(r"\n{2,}", "\n", text)
    pat = re.compile(r"(?:PROGRAMA\s+NACIONAL\s+DE\s+PROSPE[ÇC][AÃ]O\s+DE\s+PRAGAS\s+DE\s+QUARENTENA)", re.I)
    marks = [m.start() for m in pat.finditer(text)]
    if not marks or len(marks) == 1:
        return [text]
    marks.append(len(text))
    blocos = []
    for i in range(len(marks)-1):
        start = max(0, marks[i]-200); end = min(len(text), marks[i+1]+200)
        bloco = text[start:end].strip()
        if len(bloco) > 400: blocos.append(bloco)
    print(f"📄 Documento dividido em {len(blocos)} requisições.")
    return blocos

def extract_context_from_text(full_text: str) -> Dict[str, Any]:
    ctx: Dict[str, Any] = {}
    m_zona = re.search(r"Xylella\s+fastidiosa\s*\(([^)]+)\)", full_text, re.I)
    ctx["zona"] = m_zona.group(1).strip() if m_zona else "Zona Isenta"
    # DGAV
    responsavel, dgav = None, None
    m_hdr = re.search(r"Amostra(?:s|\(s\))?\s*colhida(?:s|\(s\))?\s*por\s*DGAV\s*[:\-]?\s*(.*)", full_text, re.I)
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
    # datas colheita (*) (**)
    colheita_map: Dict[str,str] = {}
    for m in re.finditer(r"(\d{1,2}/\d{1,2}/\d{4})\s*\(\s*(\*+)\s*\)", full_text):
        colheita_map[f"({m.group(2).replace(' ', '')})"] = m.group(1)
    if not colheita_map:
        m_simple = re.search(r"Data\s+de\s+colheita\s*[:\-\s]*([0-9/\-\s]+)", full_text, re.I)
        if m_simple:
            d = re.sub(r"\s+", "", m_simple.group(1))
            for key in ("(*)","(**)","(***)"): colheita_map[key] = d
    ctx["colheita_map"] = colheita_map
    ctx["default_colheita"] = next(iter(colheita_map.values()), "")
    # data envio
    m_env = re.search(r"Data\s+(?:do|de)\s+envio(?:\s+ao\s+laborat[oó]rio)?[:\-\s]*([0-9/\-\s]+)", full_text, re.I)
    ctx["data_envio"] = (re.sub(r"\s+","", m_env.group(1)) if m_env else (ctx["default_colheita"] or datetime.now().strftime("%d/%m/%Y")))
    # nº amostras declaradas
    flat = re.sub(r"\s+", " ", full_text)
    m_decl = re.search(r"N[º°]?\s*de\s*amostras(?:\s+neste\s+envio)?\s*[:\-]?\s*(\d{1,4})", flat, re.I)
    ctx["declared_samples"] = int(m_decl.group(1)) if m_decl else None
    return ctx

def parse_xylella_tables(result_json: Dict[str, Any], context: Dict[str, Any], req_id=None) -> List[Dict[str, Any]]:
    out: List[Dict[str, Any]] = []
    tables = result_json.get("analyzeResult", {}).get("tables", []) or []
    if not tables:
        print("⚠️ Nenhuma tabela encontrada."); return out

    def clean(s):
        if s is None: return ""
        return re.sub(r"\s{2,}", " ", str(s).replace("\n"," ").strip())

    for t in tables:
        cells = t.get("cells", [])
        if not cells: continue
        nc = max(c.get("columnIndex", 0) for c in cells) + 1
        nr = max(c.get("rowIndex", 0) for c in cells) + 1
        grid = [[""]*nc for _ in range(nr)]
        for c in cells:
            r, ci = c.get("rowIndex",0), c.get("columnIndex",0)
            if r<0 or ci<0: continue
            if r >= len(grid):
                grid.extend([[""]*nc for _ in range(r-len(grid)+1)])
            grid[r][ci] = clean(c.get("content",""))

        for row in grid:
            if not any(row): continue
            ref_raw = (row[0] if len(row)>0 else "").strip()
            ref = _clean_ref(ref_raw)
            if not ref or re.match(r"^\D+$", ref): continue

            hospedeiro = row[2] if len(row)>2 else ""
            obs        = row[3] if len(row)>3 else ""
            if _looks_like_natureza(hospedeiro): hospedeiro = ""

            tipo = ""
            joined = " ".join(x for x in row if isinstance(x,str))
            m_tipo = TIPO_RE.search(joined)
            if m_tipo:
                tipo = m_tipo.group(1).capitalize()
                obs = re.sub(TIPO_RE, "", obs).strip()

            datacolheita = context.get("default_colheita","")
            m_ast = re.search(r"\(\s*\*+\s*\)", joined)
            if m_ast:
                mark = re.sub(r"\s+","", m_ast.group(0))
                datacolheita = context.get("colheita_map",{}).get(mark, datacolheita)

            if obs.strip().lower() in ("simples","composta","individual"): obs = ""

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

def parse_all_requisitions(result_json: Dict[str, Any], pdf_name: str, txt_path: str) -> List[List[Dict[str, Any]]]:
    # texto global (para split e contexto)
    if txt_path and os.path.exists(txt_path):
        with open(txt_path, "r", encoding="utf-8") as f:
            full_text = f.read()
    else:
        full_text = extract_all_text(result_json)

    count, _ = detect_requisicoes(full_text)
    tables_all = result_json.get("analyzeResult", {}).get("tables", []) or []

    if count <= 1:
        context = extract_context_from_text(full_text)
        amostras = parse_xylella_tables(result_json, context, req_id=1)
        return [amostras] if amostras else []

    blocos = split_if_multiple_requisicoes(full_text)
    out: List[List[Dict[str, Any]]] = []
    for i, bloco in enumerate(blocos, start=1):
        try:
            context = extract_context_from_text(bloco)
            # filtro leve por referências detetadas no bloco
            refs_bloco = re.findall(r"\b\d{2,4}/\d{2,4}/[A-Z0-9\-]+|\b\d{7,8}\b", bloco, re.I)
            tables_filtradas = [
                t for t in tables_all
                if any(ref in " ".join(c.get("content","") for c in t.get("cells",[])) for ref in refs_bloco)
            ] or tables_all
            local = {"analyzeResult": {"tables": tables_filtradas}}
            amostras = parse_xylella_tables(local, context, req_id=i)
            if amostras: out.append(amostras)
        except Exception as e:
            print(f"❌ Erro na requisição {i}: {e}")
    return out

# ───────────────────────────────────────────────
# Escrita no TEMPLATE — 1 ficheiro por requisição
# (limpa sempre a área de dados; nunca acumula)
# ───────────────────────────────────────────────
def write_to_template(rows_per_req: List[List[Dict[str, Any]]],
                      out_base_path: str,
                      expected_count: int | None = None,
                      source_pdf: str | None = None) -> List[str]:
    template_path = Path(os.environ.get("TEMPLATE_PATH", TEMPLATE_PATH))
    if not template_path.exists():
        raise FileNotFoundError(f"TEMPLATE não encontrado: {template_path}")
    out_dir = Path(os.environ.get("OUTPUT_DIR", OUTPUT_DIR))
    out_dir.mkdir(exist_ok=True)

    sheet = "Avaliação pré registo"
    start_row = 6
    base = Path(out_base_path).stem
    out_files: List[str] = []

    # se não houver rows, ainda assim cria um ficheiro vazio (_req1.xlsx)
    effective = rows_per_req if rows_per_req else [[]]

    for idx, req in enumerate(effective, start=1):
        out_path = out_dir / f"{base}_req{idx}.xlsx"
        wb = load_workbook(template_path)
        if sheet not in wb.sheetnames:
            wb.close()
            raise KeyError(f"Folha '{sheet}' não encontrada no template.")
        ws = wb[sheet]

        # 🧹 limpeza completa da área de dados (sem tocar cabeçalhos/fórmulas)
        for r in range(start_row, ws.max_row + 1):
            for c in range(1, 13):
                ws.cell(row=r, column=c).value = None

        # ✍️ escrever linhas
        for ridx, row in enumerate(req, start=start_row):
            A = ws[f"A{ridx}"]; B = ws[f"B{ridx}"]
            rece = row.get("datarececao",""); colh = row.get("datacolheita","")
            A.value = _to_dt(rece) if _is_valid_date(rece) else rece
            B.value = _to_dt(colh) if _is_valid_date(colh) else colh
            if not _is_valid_date(rece): A.fill = RED
            if not _is_valid_date(colh): B.fill = RED

            ws[f"C{ridx}"] = row.get("referencia","")
            ws[f"D{ridx}"] = row.get("hospedeiro","")
            ws[f"E{ridx}"] = row.get("tipo","")
            ws[f"F{ridx}"] = row.get("zona","")
            ws[f"G{ridx}"] = row.get("responsavelamostra","")
            ws[f"H{ridx}"] = row.get("responsavelcolheita","")
            ws[f"I{ridx}"] = ""  # Observações
            ws[f"K{ridx}"] = row.get("procedure","XYLELLA")
            ws[f"L{ridx}"] = f"=A{ridx}+30"

            # obrigatórios A→G
            for col in ("A","B","C","D","E","F","G"):
                if not ws[f"{col}{ridx}"].value or str(ws[f"{col}{ridx}"].value).strip()=="":
                    ws[f"{col}{ridx}"].fill = RED

        processed = len(req)

        # E1:F1 — validação nº amostras
        ws.merge_cells("E1:F1")
        ws["E1"].value = f"Nº Amostras: {(expected_count if expected_count is not None else '?')} / {processed}"
        ws["E1"].font = BOLD
        ws["E1"].alignment = Alignment(horizontal="center", vertical="center")
        ws["E1"].fill = (RED if (expected_count is not None and expected_count != processed) else GREEN)

        # G1:J1 — origem do PDF
        ws.merge_cells("G1:J1")
        ws["G1"].value = f"Origem: {os.path.basename(source_pdf) if source_pdf else base}"
        ws["G1"].font = ITALIC
        ws["G1"].alignment = Alignment(horizontal="left", vertical="center")
        ws["G1"].fill = GRAY

        # K1:L1 — timestamp
        ws.merge_cells("K1:L1")
        ws["K1"].value = f"Processado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        ws["K1"].font = ITALIC
        ws["K1"].alignment = Alignment(horizontal="right", vertical="center")
        ws["K1"].fill = GRAY

        wb.save(out_path); wb.close()
        print(f"🟢 Gravado: {out_path}")
        out_files.append(str(out_path))

    return out_files

# ───────────────────────────────────────────────
# OCR + Parsing (retorna listas por requisição)
# ───────────────────────────────────────────────
def process_pdf_sync(pdf_path: str) -> List[List[Dict[str, Any]]]:
    pdf_path = str(pdf_path)
    base = Path(pdf_path).stem

    # OCR Azure
    result = azure_analyze_pdf(pdf_path)

    # Guardar texto OCR bruto (útil para debug)
    full_text = extract_all_text(result)
    debug_txt = OUTPUT_DIR / f"{base}_ocr_debug.txt"
    with open(debug_txt, "w", encoding="utf-8") as f:
        f.write(full_text)
    print(f"📝 Texto OCR bruto guardado em: {debug_txt}")

    # Parser (exato ao Colab): tabelas + contexto + split de requisições
    rows_per_req = parse_all_requisitions(result, pdf_path, str(debug_txt))
    total = sum(len(x) for x in rows_per_req)
    print(f"🏁 Concluído: {len(rows_per_req)} requisições, {total} amostras.")
    return rows_per_req


# Execução direta (opcional)
if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="Xylella Processor (Azure + Parser Colab + Template)")
    ap.add_argument("pdf", help="Caminho do PDF")
    ap.add_argument("--expected", type=int, default=None)
    args = ap.parse_args()

    rows = process_pdf_sync(args.pdf)
    base = Path(args.pdf).stem
    files = write_to_template(rows, base, expected_count=args.expected, source_pdf=Path(args.pdf).name)
    print("\n📂 Saídas:")
    for f in files: print(" -", f)
