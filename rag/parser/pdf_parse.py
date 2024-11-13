from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LTTextBoxHorizontal
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFTextExtractionNotAllowed
from pdfminer.high_level import extract_text
import logging
import warnings
logging.getLogger("pdfminer").setLevel(logging.WARNING)
warnings.filterwarnings('ignore')

class PDFparser:
    def __call__(self,file_path):
        with open(file_path, 'rb') as file:
            parser = PDFParser(file)
            doc = PDFDocument(parser)

            if not doc.is_extractable:
                raise PDFTextExtractionNotAllowed("文档不允许提取文本")

            rsrcmgr = PDFResourceManager()
            laparams = LAParams()
            device = PDFPageAggregator(rsrcmgr, laparams=laparams)
            interpreter = PDFPageInterpreter(rsrcmgr, device)

            content = ''
            for page in PDFPage.create_pages(doc):
                interpreter.process_page(page)
                layout = device.get_result()
                for element in layout:
                    if isinstance(element, LTTextBoxHorizontal):
                        content += element.get_text()
        return content