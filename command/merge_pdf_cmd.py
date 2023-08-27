import argparse
from enum import IntEnum
import logging
import os
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


class FilenamePages:
    def __init__(
        self,
        filename: str,
        ranges: Union[None, List[Tuple[int, int]]],
        front_cover: FrontBackCover = FrontBackCover.PrintBoth,
        back_cover: FrontBackCover = FrontBackCover.PrintBoth,
    ) -> None:
        self.filename: str = filename
        self.ranges: Union[None, List[Tuple[int, int]]] = ranges
        self.front_cover = front_cover
        self.back_cover = back_cover


class MergePDFCmd(BasePDFCmd):
    def __init__(self, output_filename: str, filename_pages_list: List[FilenamePages]):
        super().__init__(output_filename)

        self.filename_pages_list = filename_pages_list

    def _execute(self) -> None:
        logger.info(f"Combining pdf files to {self.output_file}.")

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
        for r in self.filename_pages_list:
            insert_at = insert_page_range(writer, insert_at, r.filename, r.ranges)

        # write result
        if len(writer.pages):
            writer.write(self.output_file)


def _construct_argparse():
    parser = argparse.ArgumentParser()
    parser.add_argument("-o", "--output", required=True, help="Output filename.")
    parser.add_argument("input_files", nargs="+", help="list of input files with optional page range. i.e. <filename>:1-3")

    return parser


def parse_page_range(str_ranges: str) -> List[Tuple[int, int]]:
    ranges = []
    for str_range in str_ranges.split(","):
        numbers = str_range.split("-")
        mn = mx = int(numbers[0])
        if len(numbers) > 1:
            mx = int(numbers[1])
        else:
            mx = mn + 1

        ranges.append((mn, mx))

    return ranges


if __name__ == "__main__":
    parser = _construct_argparse()
    args = parser.parse_args()
    if args.output:
        filename_pages_list = []
        for filename in args.input_files:
            d, fn = os.path.split(filename)
            filename_pages = FilenamePages(filename, None)
            fn_ranges = fn.split(":", maxsplit=1)
            fn = fn_ranges[0]
            if len(fn_ranges) > 1:
                ranges = parse_page_range(fn_ranges[1])
                filename_pages.ranges = ranges
            filename_pages_list.append(filename_pages)

        cmd = MergePDFCmd(args.output, filename_pages_list)
        cmd.execute()
