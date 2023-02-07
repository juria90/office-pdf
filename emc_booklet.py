from argparse import Namespace
import logging
import os
from pathlib import Path
import sys
from typing import List

from langdetect import detect  # type: ignore

from command.base_pdf_cmd import InchesToPoint  # type: ignore
from command.access_pdf_cmd import AccessPDFCmd, PrintConfig
from command.hf_pdf_cmd import KOREAN_FONT, HeaderFooterPDFCmd, TextAlign
from command.impose_pdf_cmd import ImposePDFCmd
from command.merge_content_pdf_cmd import MergeContentPDFCmd
from command.merge_pdf_cmd import MergePDFCmd, MergePDFPages
from command.word_pdf_cmd import WordPDFCmd

from init_log import init_log
from pdf_info import get_num_pages


logger = logging.getLogger(__name__)


def create_emc_booklet():
    output_dir = Path(os.path.dirname(__file__))

    # Generate family/personal contact pdfs from address.mdb
    mdb_filename = str(output_dir / "address.mdb")
    printout_configs = [
        PrintConfig(output_filename=str(output_dir / "0Pastor.pdf"), report="FAMILY-SUM", query="PMasterAnnointed", order_by="P.SID"),
        PrintConfig(output_filename=str(output_dir / "1KM-Family.pdf"), report="FAMILY-SUM", query="AddressMaster-KM", order_by="P.NAME"),
        PrintConfig(
            output_filename=str(output_dir / "2EM-Family.pdf"), report="FAMILY-SUM", query="AddressMaster-EM", order_by="SID DESC, ENAME"
        ),
        PrintConfig(output_filename=str(output_dir / "3EM-Single.pdf"), report="Single", query="EM-SINGLE", order_by="P.NAME"),
        PrintConfig(output_filename=str(output_dir / "4YG-Single.pdf"), report="Single", query="YG", order_by="P.NAME"),
    ]
    accesscmd = AccessPDFCmd(mdb_filename, printout_configs)
    accesscmd.execute()
    contact_pdf_files = accesscmd.output_files

    # Generate pdf from master word file.
    docx_filename = str(Path.home() / r"Dropbox/EMC/일반행정/2023/2023 신앙생활요람.docx")
    master_pdf_file = str(output_dir / "2023 신앙생활요람.pdf")
    wordcmd = WordPDFCmd(docx_filename, master_pdf_file)
    wordcmd.execute()

    # Combine master and contact pdfs.
    address_book_start_page_no = 29
    contact_pdf_files.insert(0, str(output_dir / "blankpage.pdf"))
    contact_pdf_files.insert(1, str(output_dir / "2023 주소록-내지.pdf"))
    contact_pdf_files.insert(2, str(output_dir / "blankpage.pdf"))
    pdf_ranges: List[MergePDFPages] = []
    pdf_ranges.append(MergePDFPages(master_pdf_file, [(0, address_book_start_page_no)]))
    pdf_ranges.append(MergePDFPages(contact_pdf_files, None))
    pdf_ranges.append(MergePDFPages(master_pdf_file, [(address_book_start_page_no + 2, 0)]))
    mergecmd = MergePDFCmd(master_pdf_file, pdf_ranges)
    mergecmd.execute()

    # Create numbering and header/footer pdf files.
    pastor_page_no = address_book_start_page_no + 3
    km_page_no = pastor_page_no + get_num_pages(contact_pdf_files[3])
    em_page_no = km_page_no + get_num_pages(contact_pdf_files[4])
    yg_page_no = em_page_no + get_num_pages(contact_pdf_files[5]) + get_num_pages(contact_pdf_files[6])
    group_title = {pastor_page_no: "사역자", km_page_no: "KM", em_page_no: "English Ministry", yg_page_no: "Youth Group"}

    statement_x = InchesToPoint(5.5)
    statement_y = InchesToPoint(8.5)

    # callback function that returns content for each page_no in pdf.
    def content_function(page_no: int):
        nonlocal group_title
        content_list = []
        # insert and blank page
        if page_no >= 2 and not (address_book_start_page_no <= page_no and page_no <= address_book_start_page_no + 2):
            content = Namespace(name="Helvetica", font_size=10, x=statement_x // 2, y=16, align=TextAlign.CENTER, text=str(page_no + 1))
            content_list.append(content)
        if page_no in group_title:
            text = group_title[page_no]
            if page_no % 2 == 1:
                x = 32
                align = TextAlign.LEFT
            else:
                x = statement_x - 32
                align = TextAlign.RIGHT
            content = Namespace(name="Helvetica", font_size=14, x=x, y=statement_y - 24, align=align, text=text)
            if detect(text) == "ko":
                content.name = KOREAN_FONT
            content_list.append(content)

        return content_list

    num_pages = get_num_pages(master_pdf_file)
    page_size = (statement_x, statement_y)
    hfcmd = HeaderFooterPDFCmd("", num_pages, page_size, content_function)
    hfcmd.execute()
    # numberingcmd = CreateNumberingPDFCmd("", num_pages, page_size, group_title)
    # numberingcmd.execute()

    # merge content of master and header/foot pdf files.
    contentcmd = MergeContentPDFCmd("", [master_pdf_file, hfcmd.output_file])
    contentcmd.execute()

    # Do saddle stitch imposition for final pdf.
    output_imposed_filename = str(output_dir / "2023 신앙생활요람-imp.pdf")
    folds = "h"  # fold horz one time to produce 1x2 saddle
    imposecmd = ImposePDFCmd(output_imposed_filename, [contentcmd.output_file], folds)
    imposecmd.execute()


if __name__ == "__main__":
    # allow interactive debugging
    if len(sys.argv) > 1 and sys.argv[1] == "--debug":
        input()
        sys.argv = sys.argv[:1] + sys.argv[2:]

    init_log()

    create_emc_booklet()
