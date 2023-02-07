import logging
from typing import List

# Use pdfimpose for imposition: pip install pdfimpose
# https://pdfimpose.readthedocs.io/en/latest/lib/saddle/
from pdfimpose.schema import saddle

from command.base_pdf_cmd import BasePDFCmd


logger = logging.getLogger(__name__)


# https://docs.reportlab.com/reportlab/userguide/ch3_fonts/#asian-font-support
KOREAN_FONT = "HYSMyeongJo-Medium"

from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont

pdfmetrics.registerFont(UnicodeCIDFont(KOREAN_FONT))


class ImposePDFCmd(BasePDFCmd):
    def __init__(self, output_filename: str, input_files: List[str], folds: str):
        super().__init__(output_filename)

        self.input_files = input_files
        self.folds = folds

    def _execute(self) -> None:
        logger.info(f"Creating imposed pdf file: {self.output_files}.")

        self._impose()

    def _impose(self):
        folds = "h"  # fold horz one time to produce 1x2 saddle
        saddle.impose(self.input_files, self.output_files, folds=folds)
