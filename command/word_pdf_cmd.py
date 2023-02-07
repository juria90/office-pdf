import logging
from typing import Tuple

from command.base_pdf_cmd import BasePDFCmd
from thirdparty.word_win32 import (
    App as WordApp,
    LineSpacingRule,
    TriState,
    WdBreakType,
    WdCollapseDirection,
    WdGoToDirection,
    WdGoToItem,
    WdHeaderFooterIndex,
    WdPageNumberAlignment,
)

logger = logging.getLogger(__name__)


class WordPDFCmd(BasePDFCmd):
    def __init__(self, docx_filename: str, output_file: str):
        super().__init__([output_file])

        self.docx_filename = docx_filename
        self.output_file = output_file

    def _execute(self):
        logger.info(f"Generating master booklet file '{self.output_file}'.")

        app = WordApp(False)
        with app.open(self.docx_filename) as doc:
            doc.print_as(self.output_file)

        app.quit()


class CreateNumberingPDFCmd(BasePDFCmd):
    def __init__(self, output_file: str, num_pages: int, page_size: Tuple[int, int], group_title: dict):
        super().__init__([output_file])

        self.output_file = output_file
        self.num_pages = num_pages
        self.page_size = page_size
        self.group_title = group_title

    def _execute(self):
        logger.info(f"Generating numbering PDF file '{self.output_file}'.")

        app = WordApp(True)
        with app.new() as doc:
            doc_obj = doc.doc
            section = doc_obj.Sections(1)

            # statement : https://python-docx.readthedocs.io/en/latest/api/shared.html#docx.shared.Length
            pagesetup = section.PageSetup
            pagesetup.PageWidth = page_size[0]
            pagesetup.PageHeight = page_size[1]

            pagesetup.HeaderDistance = 0
            pagesetup.FooterDistance = InchesToPoint(0.2)

            pagesetup.LeftMargin = InchesToPoint(0.45)
            pagesetup.RightMargin = InchesToPoint(0.45)
            pagesetup.TopMargin = InchesToPoint(0.0)
            pagesetup.BottomMargin = InchesToPoint(0.2)

            footer = section.Footers(WdHeaderFooterIndex.wdHeaderFooterPrimary)
            footer.PageNumbers.Add(WdPageNumberAlignment.wdAlignPageNumberCenter, TriState.msoTrue)
            # footer.Range.Text = "\t"

            paragraph = doc_obj.Paragraphs(1)
            paragraph.LineSpacingRule = LineSpacingRule.wdLineSpaceMultiple
            paragraph.LineSpacing = 1
            paragraph.SpaceAfterAuto = TriState.msoFalse
            paragraph.SpaceAfter = 0
            paragraph.SpaceBeforeAuto = TriState.msoFalse
            paragraph.SpaceBefore = 0

            # make num_pages pages
            for _i in range(num_pages - 1):
                myRange = doc_obj.Paragraphs(1).Range
                myRange.Collapse(WdCollapseDirection.wdCollapseEnd)
                myRange.InsertBreak(WdBreakType.wdPageBreak)

            # https://learn.microsoft.com/en-us/office/vba/api/word.selection.goto
            for page_no, text in group_title.items():
                doc_obj.ActiveWindow.Selection.GoTo(WdGoToItem.wdGoToPage, WdGoToDirection.wdGoToAbsolute, page_no + 1)
                doc_obj.ActiveWindow.Selection.Text = text
                doc_obj.ActiveWindow.Selection.Font.Size = 14

            doc.print_as(self.output_file)

        app.quit()
