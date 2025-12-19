import os
import fitz
import cv2
import pytesseract
import numpy as np
from pdf2docx import Converter
from pdf2image import convert_from_path
from docx import Document
from tqdm import tqdm

# === CONFIGURACIÓN TESSERACT ===
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
os.environ["TESSDATA_PREFIX"] = r"C:\Program Files\Tesseract-OCR"

# === CONFIGURACIÓN POPPLER ===
POPPLER_PATH = r"D:\Release-25.12.0-0\poppler-25.12.0\Library\bin"


def pdf_tiene_texto(pdf_path):
    doc = fitz.open(pdf_path)
    for page in doc:
        if page.get_text().strip():
            doc.close()
            return True
    doc.close()
    return False


def preprocesar_imagen(img):
    img = np.array(img)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    gray = cv2.medianBlur(gray, 3)
    thresh = cv2.adaptiveThreshold(
        gray, 255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY,
        31, 2
    )
    return thresh


def convertir_pdf(pdf_path, output_docx):
    print("🔍 Analizando tipo de PDF...")

    # ==========================================================
    # PDF NORMAL (TIENE TEXTO)
    # ==========================================================
    if pdf_tiene_texto(pdf_path):
        print("📄 PDF normal detectado")
        cv = Converter(pdf_path)

        total = len(cv.pages)
        print(f"📑 Páginas: {total}")

        for i in tqdm(range(total), desc="Convirtiendo páginas"):
            cv.convert(output_docx, start=i, end=i + 1)

        cv.close()
        return "PDF normal convertido correctamente"

    # ==========================================================
    # PDF ESCANEADO (OCR)
    # ==========================================================
    print("📸 PDF escaneado detectado (OCR)")
    images = convert_from_path(
        pdf_path,
        dpi=300,
        poppler_path=POPPLER_PATH
    )

    doc = Document()

    for img in tqdm(images, desc="Aplicando OCR"):
        procesada = preprocesar_imagen(img)
        texto = pytesseract.image_to_string(
            procesada,
            lang="spa",
            config="--psm 6"
        )
        doc.add_paragraph(texto)

    doc.save(output_docx)
    return "PDF escaneado convertido con OCR"
