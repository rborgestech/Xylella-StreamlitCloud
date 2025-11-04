# azure_ocr.py
# Cliente leve para OCR com fallback inteligente (Azure ou Tesseract), com cache seguro em RAM e OCR paralelo

import os
import io
import requests
import fitz  # PyMuPDF
from PIL import Image
from concurrent.futures import ThreadPoolExecutor, as_completed

# Azure configs
AZURE_KEY = os.environ.get("AZURE_KEY")
AZURE_ENDPOINT = os.environ.get("AZURE_ENDPOINT")
if not AZURE_KEY or not AZURE_ENDPOINT:
    print("\u26a0\ufe0f AVISO: AZURE_KEY ou AZURE_ENDPOINT n\u00e3o configurados \u2014 OCR Azure ser\u00e1 ignorado.")
else:
    print(f"\ud83d\udd17 Azure endpoint ativo: {AZURE_ENDPOINT}")

# Endpoint Azure
READ_URL = f"{AZURE_ENDPOINT}/computervision/imageanalysis:analyze?api-version=2023-02-01-preview&features=read"

# Cache OCR em RAM (por sess\u00e3o)
ocr_cache = {}

# ----------------------------------------------------------------------
# Verifica se o PDF tem texto embutido
# ----------------------------------------------------------------------
def has_embedded_text(pdf_path):
    with fitz.open(pdf_path) as doc:
        for page in doc:
            if page.get_text().strip():
                return True
    return False

# ----------------------------------------------------------------------
# Converte PDF em imagens (PyMuPDF, eficiente)
# ----------------------------------------------------------------------
def pdf_to_images(pdf_path):
    images = []
    with fitz.open(pdf_path) as doc:
        for page in doc:
            pix = page.get_pixmap(dpi=200)
            img = Image.open(io.BytesIO(pix.tobytes("png")))
            images.append(img)
    return images

# ----------------------------------------------------------------------
# Envia imagem para Azure OCR (com cache)
# ----------------------------------------------------------------------
def extract_text_from_image_azure_bytes(img_bytes: bytes, page_idx: int = 0):
    key = f"page_{page_idx}_{hash(img_bytes)}"
    if key in ocr_cache:
        return ocr_cache[key]

    headers = {
        "Ocp-Apim-Subscription-Key": AZURE_KEY,
        "Content-Type": "application/octet-stream"
    }
    response = requests.post(READ_URL, headers=headers, data=img_bytes)
    if response.status_code not in (200, 202):
        raise RuntimeError(f"Erro no envio OCR: {response.status_code} - {response.text}")

    result = response.json()
    ocr_cache[key] = result
    return result

# ----------------------------------------------------------------------
# Normaliza JSON Azure
# ----------------------------------------------------------------------
def get_analysis_result_azure(result_json):
    if "analyzeResult" in result_json:
        return result_json
    return {"analyzeResult": result_json}

# ----------------------------------------------------------------------
# OCR paralelo Azure (ThreadPool)
# ----------------------------------------------------------------------
def ocr_parallel_azure(images):
    text_total = ""
    with ThreadPoolExecutor(max_workers=4) as executor:
        futures = []
        for idx, img in enumerate(images, start=1):
            buf = io.BytesIO()
            img.save(buf, format="JPEG")
            img_bytes = buf.getvalue()
            futures.append(executor.submit(extract_text_from_image_azure_bytes, img_bytes, idx))

        for i, future in enumerate(as_completed(futures), start=1):
            try:
                result = future.result()
                data = get_analysis_result_azure(result)
                for block in data.get("analyzeResult", {}).get("readResult", []):
                    for line in block.get("lines", []):
                        text_total += line.get("text", "") + "\n"
            except Exception as e:
                print(f"\u26a0\ufe0f Erro OCR p\u00e1gina {i}: {e}")
    return text_total

# ----------------------------------------------------------------------
# Extra\u00e7\u00e3o universal de texto
# ----------------------------------------------------------------------
def extract_all_text(pdf_path):
    import pytesseract

    # Verifica se tem texto embutido
    if has_embedded_text(pdf_path):
        print("\ud83d\udcc4 PDF com texto embutido \u2014 extra\u00e7\u00e3o direta.")
        text_total = ""
        with fitz.open(pdf_path) as doc:
            for page in doc:
                text_total += page.get_text() + "\n"
        return text_total

    print("\ud83d\udd0d PDF sem texto \u2014 a aplicar OCR.")
    images = pdf_to_images(pdf_path)

    if AZURE_KEY and AZURE_ENDPOINT:
        return ocr_parallel_azure(images)
    else:
        print("\u26a0\ufe0f OCR Azure indispon\u00edvel \u2014 a usar OCR local.")
        text_total = ""
        for img in images:
            text_total += pytesseract.image_to_string(img, lang="por")
        return text_total
