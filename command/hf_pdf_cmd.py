from enum import IntEnum
import logging
from typing import Callable, Tuple

from reportlab.pdfgen import canvas  # type:ignore

from command.base_pdf_cmd import BasePDFCmd


logger = logging.getLogger(__name__)


# https://docs.reportlab.com/reportlab/userguide/ch3_fonts/#asian-font-support
KOREAN_FONT = "HYSMyeongJo-Medium"

from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont

pdfmetrics.registerFont(UnicodeCIDFont(KOREAN_FONT))


class TextAlign(IntEnum):
    LEFT = 0
    CENTER = 1
    RIGHT = 2


class HeaderFooterPDFCmd(BasePDFCmd):
    def __init__(self, output_filename: str, num_pages: int, page_size: Tuple[int, int], content_function: Callable):
        super().__init__(output_filename)

        self.num_pages = num_pages
        self.page_size = page_size
        self.content_function = content_function

    def _execute(self) -> None:
        logger.info(f"Creating page numbers and titles in PDF file: {self.output_file}.")

        self.create_pdf_text_pages()

    def create_pdf_text_pages(self):
        c = canvas.Canvas(self.output_file)
        # logger.info(f"getAvailableFonts: {c.getAvailableFonts()}")
        for page_no in range(self.num_pages):
            c.setPageSize(self.page_size)
            content_list = self.content_function(page_no)
            if isinstance(content_list, list) and len(content_list) > 0:
                for content in content_list:
                    c.setFont(content.name, content.font_size)  # choose your font name and font size
                    if content.align == TextAlign.CENTER:
                        c.drawCentredString(content.x, content.y, content.text)
                    elif content.align == TextAlign.RIGHT:
                        c.drawRightString(content.x, content.y, content.text)
                    else:
                        c.drawString(content.x, content.y, content.text)
            c.showPage()
        c.save()
