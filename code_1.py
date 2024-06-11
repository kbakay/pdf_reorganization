from pdf2docx import Converter
from docx import Document
import os

# PDF'i Word dosyasına dönüştürme
def pdf_to_docx(pdf_path, docx_path):
    cv = Converter(pdf_path)
    cv.convert(docx_path)
    cv.close()

# Ana işlem
def main():
    # PDF dosyalarının bulunduğu klasör
    pdf_folder = "A"
    # Word dosyalarının kaydedileceği klasör
    docx_folder = "1"

    # Kontrol etme ve oluşturma
    if not os.path.exists(docx_folder):
        os.makedirs(docx_folder)

    # PDF dosyalarını Word dosyasına dönüştürme ve kaydetme
    for filename in os.listdir(pdf_folder):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(pdf_folder, filename)
            docx_path = os.path.join(docx_folder, os.path.splitext(filename)[0] + ".docx")
            pdf_to_docx(pdf_path, docx_path)

            # Word dosyasını açma ve işlemleri yapma
            doc = Document(docx_path)
            # Buraya tablo birleştirme ve diğer işlemler eklenebilir

            # Sonuçları kaydetme
            doc.save(os.path.join(docx_folder, os.path.splitext(filename)[0] + "_updated.docx"))


if __name__ == "__main__":
    main()
