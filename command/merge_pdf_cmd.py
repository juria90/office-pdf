from enum import IntEnum
import logging
from typing import List, Tuple, Union

from PyPDF2 import PdfWriter

from command.base_pdf_cmd import BasePDFCmd
from pdf_info import open_pdfreader


logger = logging.getLogger(__name__)


class FrontBackCover(IntEnum):
    PrintBoth = 0
    PrintFront = 1
    PrintBack = 2
    BlankBoth = 3


class MergePDFPages:
    def __init__(
        self,
        filenames: Union[str, List[str]],
        ranges: Union[None, List[Tuple[int, int]]],
        front_cover: FrontBackCover = FrontBackCover.PrintBoth,
        back_cover: FrontBackCover = FrontBackCover.PrintBoth,
    ) -> None:
        self.filenames: Union[str, List[str]] = filenames
        self.ranges: Union[None, List[Tuple[int, int]]] = ranges
        self.front_cover = front_cover
        self.back_cover = back_cover


class MergePDFCmd(BasePDFCmd):
    def __init__(self, output_filename: str, pdf_ranges: List[MergePDFPages]):
        super().__init__(output_filename)

        self.pdf_ranges = pdf_ranges

    def _execute(self) -> None:
        logger.info(f"Combining pdf files to {self.output_files}.")

        writer = PdfWriter()

        def insert_page_range(writer: PdfWriter, insert_at: int, filename: str, ranges: Union[None, List[Tuple[int, int]]]):
            with open_pdfreader(filename) as pdf:
                num_pages = len(pdf.pages)
                if insert_at == -1:
                    insert_at = len(writer.pages)
                if ranges is None:
                    for i in range(num_pages):
                        writer.insert_page(pdf.pages[i], insert_at)
                        insert_at += 1
                elif isinstance(ranges, list):
                    for rng in ranges:
                        mx = rng[1]
                        if mx <= 0:
                            mx = num_pages + mx
                        for i in range(rng[0], mx):
                            writer.insert_page(pdf.pages[i], insert_at)
                            insert_at += 1

            return insert_at

        insert_at = 0
        for r in self.pdf_ranges:
            if isinstance(r.filenames, list):
                for filename in r.filenames:
                    insert_at = insert_page_range(writer, insert_at, filename, r.ranges)
            elif isinstance(r.filenames, str):
                filename = r.filenames
                insert_at = insert_page_range(writer, insert_at, filename, r.ranges)

        # write result
        if len(writer.pages):
            writer.write(self.output_files)
