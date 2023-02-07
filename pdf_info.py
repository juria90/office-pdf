from contextlib import contextmanager
import logging

from PyPDF2 import PdfReader


logger = logging.getLogger(__name__)


@contextmanager
def open_pdfreader(filename: str):
    reader = PdfReader(filename, strict=False)
    yield reader
    # reader.


def get_num_pages(filename: str) -> int:
    num_pages = 0
    with open_pdfreader(filename) as reader:
        num_pages = len(reader.pages)

    return num_pages
