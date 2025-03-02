from pdf2docx import Converter
from PySide6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QLabel)
import sys


class PDFtoDOCXConverterApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF to DOCX Converter")
        self.setGeometry(100, 100, 400, 200)

        layout = QVBoxLayout()

        self.label = QLabel("Select a PDF file to convert")
        layout.addWidget(self.label)

        self.select_button = QPushButton("Select PDF")
        self.select_button.clicked.connect(self.select_pdf)
        layout.addWidget(self.select_button)

        self.convert_button = QPushButton("Convert to DOCX")
        self.convert_button.clicked.connect(self.convert_pdf_to_docx)
        layout.addWidget(self.convert_button)

        self.setLayout(layout)

        self.pdf_file = ""
        self.docx_file = ""

    def select_pdf(self):
        self.pdf_file, _ = QFileDialog.getOpenFileName(self, "Open PDF File", "", "PDF Files (*.pdf)")
        if self.pdf_file:
            self.label.setText(f"Selected PDF: {self.pdf_file}")
            self.docx_file = self.pdf_file.replace(".pdf", ".docx")

    def convert_pdf_to_docx(self):
        if not self.pdf_file:
            self.label.setText("Please select a PDF file first!")
            return
        try:
            cv = Converter(self.pdf_file)
            cv.convert(self.docx_file, start=0, end=None)
            cv.close()
            self.label.setText(f"Conversion successful: {self.docx_file}")
        except Exception as e:
            self.label.setText(f"Error during conversion: {e}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PDFtoDOCXConverterApp()
    window.show()
    sys.exit(app.exec())
