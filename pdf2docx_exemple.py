from pdf2docx import Converter


def pdf_to_docx(pdf_file, docx_file):
    try:
        cv = Converter(pdf_file)
        cv.convert(docx_file, start=0, end=None)
        cv.close()
        print(f"Conversion successful: {docx_file}")
    except Exception as e:
        print(f"Error during conversion: {e}")


if __name__ == "__main__":
    pdf_file = "programari_html.pdf" # Specifică calea către fișierul PDF
    docx_file = "document.docx"  # Specifică calea pentru fișierul DOCX de ieșire
    pdf_to_docx(pdf_file, docx_file)
