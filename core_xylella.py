# -*- coding: utf-8 -*-
"""
CORE_XYLELLA.PY — Motor principal do processamento Xylella

Responsável por:
 - Executar OCR via Azure (página a página)
 - Extrair tabelas e texto global
 - Analisar blocos de requisição
 - Exportar resultados para o TEMPLATE Excel

Funções principais expostas:
    • process_pdf_sync(pdf_path) → rows
    • write_to_template(rows, pdf_name)
"""

from __future__ import annotations
import os, re
import tempfile
from pathlib import Path
from datetime import datetime
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill
import pytesseract
from azure_ocr import pdf_to_images, extract_text_from_image_azure, get_analysis_result_azure




# ── Caminhos globais ────────────────────────────────────────────────
BASE_DIR = Path(__file__).parent
OUTPUT_DIR = BASE_DIR / "Output"
TEMPLATE_PATH = BASE_DIR / "TEMPLATE_PXf_SGSLABIP1056.xlsx"

# Garantir que a pasta Output existe
OUTPUT_DIR.mkdir(exist_ok=True)


# ───────────────────────────────────────────────────────────────
#  Funções auxiliares externas (OCR Azure)
# ───────────────────────────────────────────────────────────────
# Estas devem estar definidas noutro módulo (azure_ocr.py):
#  • pdf_to_images(pdf_path)
#  • extract_text_from_image_azure(image_path)
#  • get_analysis_result_azure(result_url)
#  • extract_all_text(result_json)
#  • validate_plants(rows)

# ----------------------------------------------------------------
#  Funções utilitárias e normalização
# ----------------------------------------------------------------

def normalize_dedup(rows):
    """Remove duplicados e normaliza nomes."""
    cleaned = []
    seen = set()
    for r in rows:
        if not r.get("referencia"):
            continue
        r["hospedeiro"] = re.sub(r"[%\.,;:]+$", "", str(r.get("hospedeiro", ""))).strip()
        r["hospedeiro"] = re.sub(r"\s+", " ", r["hospedeiro"])
        key = (r["referencia"], r["hospedeiro"].lower(), r["tipo"].lower())
        if key in seen:
            continue
        seen.add(key)
        cleaned.append(r)
    return cleaned


# ----------------------------------------------------------------
#  Escrever resultado no TEMPLATE Excel
# ----------------------------------------------------------------
def write_to_template(ocr_rows, out_base_path, expected_count=None, source_pdf=None):
    """
    Escreve as requisições no TEMPLATE_PXF_SGSLABIP1056.xlsx
    mantendo fórmulas, validações e formatação SGS.
    """
    from openpyxl import load_workbook
    import shutil

    template_path = Path(os.environ["TEMPLATE_PATH"])
    if not template_path.exists():
        raise FileNotFoundError(f"TEMPLATE não encontrado: {template_path}")

    out_files = []
    start_row = 6
    sheet_name = "Avaliação pré registo"

    for idx, req_rows in enumerate(ocr_rows, start=1):
        out_path = Path(f"{out_base_path}_req{idx}.xlsx")
        shutil.copy(template_path, out_path)

        wb = load_workbook(out_path)
        ws = wb[sheet_name]

        # escreve as linhas extraídas sem tocar nas fórmulas existentes
        for i, row in enumerate(req_rows, start=start_row):
            for j, value in enumerate(row, start=1):
                ws.cell(row=i, column=j).value = value

        # validação opcional: destaca se o nº de amostras for diferente do esperado
        if expected_count and len(req_rows) != expected_count:
            from openpyxl.styles import PatternFill
            fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            ws["E1"].fill = fill
            ws["E1"].value = f"Atenção: {len(req_rows)} amostras, esperado {expected_count}"

        # metadata
        if source_pdf:
            ws["F1"].value = f"Origem: {source_pdf}"

        wb.save(out_path)
        print(f"🟢 Gravado com sucesso: {out_path}")
        out_files.append(out_path)

    return out_files


# ----------------------------------------------------------------
#  Função principal: Processamento do PDF (síncrono)
# ----------------------------------------------------------------
def process_pdf_sync(pdf_path):
    """
    Extrai texto de um PDF usando OCR Azure (se configurado) ou OCR local como fallback.
    Divide automaticamente o texto por requisições e devolve listas de linhas.
    """
    import tempfile
    from azure_ocr import (
        pdf_to_images,
        extract_text_from_image_azure,
        get_analysis_result_azure,
    )

    pdf_path = str(pdf_path)
    text_total = ""

    # 1️⃣ Converte o PDF em imagens (usando PyMuPDF via azure_ocr)
    try:
        images = pdf_to_images(pdf_path)
        print(f"📄 PDF convertido em {len(images)} imagem(ns).")
    except Exception as e:
        raise RuntimeError(f"Falha ao converter PDF em imagens: {e}")

    # 2️⃣ Tenta OCR (Azure ou local) página a página
    for idx, img in enumerate(images, start=1):
        tmp_path = os.path.join(tempfile.gettempdir(), f"page_{idx}.png")
        img.save(tmp_path, "PNG")

        try:
            # tenta Azure
            result = extract_text_from_image_azure(tmp_path)
            data = get_analysis_result_azure(result)

            for block in data.get("analyzeResult", {}).get("readResult", []):
                for line in block.get("lines", []):
                    text_total += line.get("text", "") + "\n"

        except Exception as e:
            # fallback para OCR local
            print(f"⚠️ Erro no OCR Azure (página {idx}: {e}) — a usar Tesseract local.")
            import pytesseract
            text_total += pytesseract.image_to_string(img) + "\n"

    if not text_total.strip():
        raise RuntimeError(f"Não foi possível extrair texto de {os.path.basename(pdf_path)}")

    # 3️⃣ Divide o texto em requisições (usando o parser existente)
    blocos = split_if_multiple_requisicoes(text_total)
    todas = []
    for bloco in blocos:
        linhas = parse_with_regex(bloco)
        if linhas:
            todas.append(linhas)

    return todas


# ----------------------------------------------------------------
#  Parser Xylella (secção 5)
# ----------------------------------------------------------------

NATUREZA_KEYWORDS = [
    "ramos", "folhas", "ramosefolhas", "ramosc/folhas",
    "material", "materialherbalho", "materialherbário", "materialherbalo",
    "natureza", "insetos", "sementes", "solo"
]

REF_SLASH_RE = re.compile(
    r"\b\d{1,4}\s*/\s*[A-Za-z0-9]{1,10}(?:\s*/\s*[A-Za-z0-9\-]{1,12})+\b",
    re.I
)
REF_NUM_RE = re.compile(r"\b\d{7,8}\b")
TIPO_RE = re.compile(r"\b(Composta|Simples|Individual)\b", re.I)
ONLY_DATE_RE = re.compile(r"^\d{1,2}/\d{1,2}/\d{4}$")
NATUREZA_RE = re.compile(r"\bpartes?\s+de?\s*vegetais?\b|\bnatureza\s+da\s+amostra\b", re.I)


def _clean_ref(raw: str) -> str:
    s = raw.strip()
    s = re.sub(r"[\s\t]+", "", s)
    s = re.sub(r"[|,;]+", "/", s)
    s = re.sub(r"[^\w/]", "", s)
    s = re.sub(r"/{2,}", "/", s)
    return s.upper()


def _is_natureza_line(s: str) -> bool:
    t = re.sub(r"\s+", " ", s.strip().lower())
    if NATUREZA_RE.search(t):
        return True
    return any(k in t.replace(" ", "") for k in NATUREZA_KEYWORDS)


def _merge_host(lines, j):
    def ok_host(txt):
        return (re.search(r"[A-Za-zÀ-ÿ]", txt)
                and not _is_natureza_line(txt)
                and not TIPO_RE.search(txt))
    a = lines[j].strip()
    if not ok_host(a):
        return "", 0
    if j + 1 < len(lines):
        b = lines[j+1].strip()
        if ok_host(b) and re.match(r"^[A-Za-zÀ-ÿ\.\-]+$", b):
            return f"{a} {b}".strip(), 2
    return a, 1


# ───────────────────────────────────────────────────────────────
#  PARSER DE TABELAS COM CONTEXTO GLOBAL
# ───────────────────────────────────────────────────────────────
def parse_xylella_tables_from_text(full_text: str, context: dict, req_id=None):
    out = []
    lines = [l.strip() for l in full_text.splitlines() if l.strip()]
    n = len(lines)
    i = 0

    while i < n:
        line = lines[i]
        if (ONLY_DATE_RE.match(line) or
            re.search(r"^(Data\s+(de|do)\s+(envio|colheita)|N[º°]\s*de\s*amostras|PROGRAMA\s+DE|Refer|Observa|SGS)", line, re.I)):
            i += 1
            continue

        mref = REF_SLASH_RE.search(line) or REF_NUM_RE.search(line)
        if not mref:
            i += 1
            continue

        ref = _clean_ref(mref.group(0))

        if i + 1 < n:
            nxt = lines[i+1].strip()
            if re.match(r"^(?:EDM|LVT|ALG|NRT|DGAV)[\w\-]*/\d{2,4}\b", nxt, re.I):
                ref = _clean_ref(ref + "/" + nxt)
                i += 1

        if ONLY_DATE_RE.fullmatch(ref):
            i += 1
            continue

        hospedeiro = ""
        tipo = ""
        datacolheita = context.get("default_colheita", "")
        j = i + 1
        end = min(n, i + 8)

        while j < end:
            ln = lines[j]
            mt = TIPO_RE.search(ln)
            if mt and not tipo:
                tipo = mt.group(1).capitalize()

            for look_ahead in range(0, 3):
                if j + look_ahead < n:
                    ln_date = lines[j + look_ahead]
                    mast = re.search(r"\(\s*(\*+)\s*\)", ln_date)
                    if mast and context.get("colheita_map"):
                        mark = "(" + mast.group(1).replace(" ", "") + ")"
                        datacolheita = context["colheita_map"].get(mark, datacolheita)
                        break

            if not hospedeiro and not _is_natureza_line(ln) and not TIPO_RE.search(ln):
                cand, consumed = _merge_host(lines, j)
                if cand:
                    hospedeiro = re.sub(r"\s{2,}", " ", cand).strip()
                    j += consumed
                    for k in range(j, min(n, j + 2)):
                        mt2 = TIPO_RE.search(lines[k])
                        if mt2:
                            tipo = mt2.group(1).capitalize()
                    break
            j += 1

        if TIPO_RE.fullmatch(hospedeiro):
            hospedeiro = ""
        if re.match(r"^(?:EDM|LVT|ALG|NRT|DGAV)[\w\-]*/\d{2,4}\b", (hospedeiro or ""), re.I):
            hospedeiro = ""

        out.append({
            "requisicao_id": req_id,
            "datarececao": context.get("data_envio", ""),
            "datacolheita": datacolheita,
            "referencia": ref,
            "hospedeiro": hospedeiro,
            "tipo": tipo,
            "zona": context.get("zona", ""),
            "responsavelamostra": context.get("dgav", ""),
            "responsavelcolheita": "",
            "observacoes": "",
            "procedure": "XYLELLA",
            "datarequerido": context.get("data_envio", ""),
            "Score": ""
        })
        i = max(i + 1, j)

    print(f"✅ {len(out)} amostras extraídas (req_id={req_id}) do texto OCR.")
    return out


# ----------------------------------------------------------------
#  SPLIT E CONTEXTO GLOBAL
# ----------------------------------------------------------------

def split_if_multiple_requisicoes(full_text: str):
    import re
    text = re.sub(r"[ \t]+", " ", full_text)
    text = re.sub(r"\n{2,}", "\n", text)
    pattern = re.compile(
        r"(?:(?:^|\n)\s*(?:PROGRAMA\s+DE\s+PROSPE|Amostra\s+colhida\s+por\s+DGAV|Refer[eê]ncia\s+da\s+amostra))",
        re.IGNORECASE
    )
    matches = list(pattern.finditer(text))
    if not matches:
        print("🔍 Nenhum marcador de nova requisição detetado.")
        return [text]
    positions = []
    last_pos = -9999
    for m in matches:
        if m.start() - last_pos > 1200:
            positions.append(m.start())
            last_pos = m.start()
    if len(positions) == 1:
        print(f"🔍 Detetada 1 requisição (posições: {positions})")
        return [text]
    blocos = []
    for i, start in enumerate(positions):
        end = positions[i + 1] if i + 1 < len(positions) else len(text)
        blocos.append(text[start:end].strip())
    print(f"📄 Documento dividido em {len(blocos)} requisições distintas.")
    return blocos


def extract_context_from_text(full_text: str):
    import re
    from datetime import datetime
    context = {}
    m_zona = re.search(r"Xylella\s+fastidiosa\s*\(([^)]+)\)", full_text, re.I)
    context["zona"] = m_zona.group(1).strip() if m_zona else "Zona Isenta"

    responsavel = None
    m_base = re.search(
        r"Amostra.{0,40}?colhida.{0,15}?por.{0,10}?DGAV\s*[:\-]?",
        full_text, re.IGNORECASE | re.DOTALL,
    )
    if m_base:
        sub = full_text[m_base.end():]
        for ln in sub.strip().splitlines():
            ln = ln.strip()
            if not ln or re.search(r"Data|Refer", ln, re.I):
                break
            responsavel = re.sub(r"[\.:\-;,]+$", "", ln.strip())
            break
    context["responsavel_colheita"] = responsavel
    if responsavel and re.match(r"^DGAV\b", responsavel, re.I):
        context["dgav"] = responsavel
    elif responsavel:
        context["dgav"] = f"DGAV {responsavel}"
    else:
        m_dgav = re.search(r"DGAV\s+[A-ZÀ-ÿ\- ]{2,30}", full_text, re.I)
        context["dgav"] = m_dgav.group(0).strip() if m_dgav else "DGAV"

    colheita_map = {}
    text_norm = re.sub(r"\s+(?:e|ou)\s+", " ", full_text)
    text_norm = text_norm.replace(",", " ")
    for m in re.finditer(r"(\d{1,2}/\d{1,2}/\d{4})\s*\(\s*(\*+)\s*\)", text_norm):
        mark = "(" + m.group(2).replace(" ", "") + ")"
        colheita_map[mark] = m.group(1)
    if not colheita_map:
        m_simple = re.search(r"Data\s+de\s+colheita.*?([\d/]{8,10})", full_text, re.I)
        if m_simple:
            d = m_simple.group(1)
            colheita_map["(*)"] = d
            colheita_map["(**)"] = d
    context["colheita_map"] = colheita_map
    context["default_colheita"] = next(iter(colheita_map.values()), "")
    m_envio = re.search(r"Data\s+(?:do|de)\s+envio.*?([\d/]{8,10})", full_text, re.I)
    context["data_envio"] = m_envio.group(1) if m_envio else context["default_colheita"] or datetime.now().strftime("%d/%m/%Y")

    print(f"🌍 Zona de origem: {context['zona']}")
    print(f"👤 Responsável DGAV: {context['dgav']}")
    print(f"👷 Responsável pela colheita: {context['responsavel_colheita'] or '(não identificado)'}")
    print(f"📅 Datas de colheita: {colheita_map or '(nenhuma)'} (padrão: {context['default_colheita'] or 'nenhuma'})")
    print(f"📬 Data do envio ao laboratório: {context['data_envio']}")
    return context


def parse_xylella_from_result(result_json, pdf_path, txt_path=None):
    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
    if txt_path and os.path.exists(txt_path):
        with open(txt_path, "r", encoding="utf-8") as f:
            full_text = f.read()
        print(f"📝 Contexto extraído a partir de {os.path.basename(txt_path)}")
    else:
        print("⚠️ Ficheiro texto não encontrado — fallback.")
        first_page_text = "\n".join(line.get("content", "") for line in result_json.get("analyzeResult", {}).get("pages", [])[0].get("lines", []))
        full_text = first_page_text

    blocos = split_if_multiple_requisicoes(full_text)
    num_blocks = len(blocos)
    print(f"📄 Documento contém {num_blocks} bloco(s) de requisição.")

    total_validos, total_ignorados = 0, 0
    all_samples = []

    for i, bloco in enumerate(blocos, start=1):
        print(f"\n🔹 A processar requisição {i}/{num_blocks}...")
        context = extract_context_from_text(bloco)
        amostras = parse_xylella_tables_from_text(bloco, context, req_id=i)
        if not amostras:
            print(f"⚠️ Requisição {i} ignorada — sem referências válidas.")
            total_ignorados += 1
            continue
        total_validos += 1
        all_samples.extend(amostras)
        output_name = f"{base_name}_req{i}.xlsx" if num_blocks > 1 else f"{base_name}.xlsx"
        output_path = os.path.join(OUTPUT_DIR, output_name)
        wb = Workbook(); ws = wb.active
        ws.append(list(amostras[0].keys()))
        for a in amostras: ws.append(list(a.values()))
        wb.save(output_path)
        print(f"✅ Exportado: {output_path}")

    print("\n📊 Resumo de processamento:")
    print(f"   • Total de blocos: {num_blocks}")
    print(f"   • Válidos: {total_validos}")
    print(f"   • Ignorados: {total_ignorados}")
    print(f"✅ Total global: {len(all_samples)} amostras extraídas.")
    print(f"📂 Ficheiros guardados em: {OUTPUT_DIR}")

    return all_samples, num_blocks







