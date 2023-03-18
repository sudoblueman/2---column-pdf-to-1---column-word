#By changing 2-column pdf to 1-column  word, and can be converted to epub by Calibre. 
#Inspired by the isssue : https://github.com/yihong0618/bilingual_book_maker/issues/49

!pip install pdfminer
!pip install pdfminer.six
!pip install python-docx
import os
import re
from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextContainer, LTTextLine, LAParams
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

def read_pdf_file(pdf_file_path):
    text_content = ""
    with open(pdf_file_path, 'rb') as pdf_file:
        # 使用 LAParams 設置參數，使得 PDFMiner 可以識別文本內容
        laparams = LAParams()
        for page_layout in extract_pages(pdf_file, laparams=laparams):
            for element in page_layout:
                # 只選擇包含文本的元素
                if isinstance(element, (LTTextContainer, LTTextLine)):
                    # 使用正則表達式刪除或替換掉無效字符
                    pattern = re.compile('[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]')
                    text = pattern.sub('', element.get_text())
                    text_content += text + "\n"
    return text_content

def create_docx_file(docx_file_path, text_content):
    document = Document()
    paragraphs = text_content.split('\n\n')
    for paragraph in paragraphs:
        if paragraph.strip():
            doc_paragraph = document.add_paragraph(paragraph.strip())
            doc_paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    document.save(docx_file_path)

pdf_folder = '/content/input'
docx_folder = '/content/output'

if not os.path.exists(docx_folder):
    os.makedirs(docx_folder)

for filename in os.listdir(pdf_folder):
    if filename.endswith('.pdf'):
        pdf_file_path = os.path.join(pdf_folder, filename)
        text_content = read_pdf_file(pdf_file_path)
        docx_file_path = os.path.join(docx_folder, filename.replace('.pdf', '.docx'))
        create_docx_file(docx_file_path, text_content)
