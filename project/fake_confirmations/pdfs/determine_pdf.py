from PyPDF2 import PdfReader
import fitz
from pdfminer.high_level import extract_text
import re
import os

file_path = os.path.dirname(__file__)
os.chdir(file_path)


class PurchaseOrderReader:
    def __init__(self, pdf):
        self.pdf = pdf
        self.pypdf_text = ""
        self.pymupdf_text = ""

    def pdfminer_data_extracter(self):
        print("\nPDFMINER\n___________________________________________________")
        self.pdfminer_text = repr(extract_text(self.pdf))
        print(self.pdfminer_text)
        return self.pdfminer_text

    def pypdf2_data_extractor(self):
        print("\nPYPDF2\n___________________________________________________")
        reader = PdfReader(self.pdf)
        number_of_pages = len(reader.pages)
        for pagenumber in range(number_of_pages):
            page = reader.pages[pagenumber]
            self.pypdf_text = "{}\n{}".format(
                self.pypdf_text, repr(page.extract_text(space_width=200))
            )
        print(self.pypdf_text)
        return self.pypdf_text

    def pymupdf_data_extractor(self):
        print("\nPYMUPDF\n___________________________________________________")
        with fitz.open(self.pdf) as doc:  # open document
            for page in doc:
                self.pymupdf_text = "{}\n{}".format(
                    self.pymupdf_text, repr(page.get_text())
                )
        print(self.pymupdf_text)
        return self.pymupdf_text

    def match_patterns(self):
        key = ["PDFMINER", "PYPDF2", "PYMUPDF"]
        values = [self.pdfminer_text, self.pypdf_text, self.pymupdf_text]
        pdf_module_dict = dict(zip(key, values))
        for module in pdf_module_dict:
            print(
                "\n{}\n____________________________________________________".format(
                    module
                )
            )
            for orderline in range(1, 10):
                pattern = rf"((\D[n]{orderline}[.]\D[n]123)(.*?)(\D[n]€\s{{1,50}}\D[n][0-9,.]{{4,9}}\D[n]€))"
                matches = re.search(pattern, pdf_module_dict[module])
                if matches:
                    print(matches.group())
                else:
                    pattern = rf"((\D[n]{orderline}[.] 123)(.*?)(€\s{{1,50}}\D[n]))"
                    matches = re.search(pattern, pdf_module_dict[module])
                    if matches:
                        print(matches.group())
                    else:
                        print("No match found for both patterns.")


def main(pdf):
    extractor = PurchaseOrderReader(pdf)
    extractor.pdfminer_data_extracter()
    extractor.pypdf2_data_extractor()
    extractor.pymupdf_data_extractor()
    extractor.match_patterns()


if __name__ == "__main__":
    pdf = f"{file_path}/fake_confirmations/blog/Confirmation_2410182_9249895.pdf"
    main(pdf)
