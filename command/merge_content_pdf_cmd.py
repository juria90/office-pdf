from contextlib import ExitStack
import logging
from typing import List

from PyPDF2 import PdfWriter

from command.base_pdf_cmd import BasePDFCmd
from pdf_info import open_pdfreader


logger = logging.getLogger(__name__)


class MergeContentPDFCmd(BasePDFCmd):
    def __init__(self, output_filename: str, input_files: List[str]):
        super().__init__(output_filename)

        self.input_files = input_files

    def _execute(self) -> None:
        logger.info(f"Merging contents to pdf file: {self.output_files}.")

        self.merge_pdf_content()

    def merge_pdf_content(self):
        writer = PdfWriter()

        with ExitStack() as stack:
            readers = [stack.enter_context(open_pdfreader(filename)) for filename in self.input_files]
            readers0 = readers[0]
            readers1_ = readers[1:]
            num_pages = len(readers0.pages)

            # iterarte pages
            for page_no in range(num_pages):
                main_page = readers0.pages[page_no]

                # merge content in content_page to the main_page
                for r in readers1_:
                    content_page = r.pages[page_no]
                    main_page.merge_page(content_page)
                writer.add_page(main_page)

        # write result
        if len(writer.pages):
            writer.write(self.output_files)
