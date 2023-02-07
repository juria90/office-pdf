from enum import IntEnum
import os

import pythoncom  # type: ignore
import win32com.client  # type: ignore

from process_exists import process_exists


# https://learn.microsoft.com/en-us/office/vba/api/overview/word


class TriState(IntEnum):
    # https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.core.msotristate?view=office-pia
    msoFalse = 0
    msoMixed = -2
    msoTrue = -1


def TriStateToBool(value):
    return value != TriState.msoFalse


def BoolToTriState(value):
    return int(-1 if value else 0)


class WdBreakType(IntEnum):
    # https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.wdbreaktype
    wdSectionBreakNextPage = 2
    wdSectionBreakContinuous = 3
    wdSectionBreakEvenPage = 4
    wdSectionBreakOddPage = 5
    wdLineBreak = 6
    wdPageBreak = 7
    wdColumnBreak = 8
    wdLineBreakClearLeft = 9
    wdLineBreakClearRight = 10
    wdTextWrappingBreak = 11


class WdCollapseDirection(IntEnum):
    # https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.wdcollapsedirection
    wdCollapseEnd = 0
    wdCollapseStart = 1


class WdExportFormat(IntEnum):
    # https://learn.microsoft.com/en-us/office/vba/api/word.wdexportformat
    wdExportFormatPDF = 17
    wdExportFormatXPS = 18


class WdGoToDirection(IntEnum):
    # https://learn.microsoft.com/en-us/office/vba/api/word.wdgotodirection
    wdGoToLast = -1
    wdGoToFirst = 1
    wdGoToNext = 2
    wdGoToPrevious = 3
    wdGoToAbsolute = 1
    wdGoToRelative = 2


class WdGoToItem(IntEnum):
    # https://learn.microsoft.com/en-us/office/vba/api/word.wdgotoitem
    wdGoToBookmark = -1
    wdGoToSection = 0
    wdGoToPage = 1
    wdGoToTable = 2
    wdGoToLine = 3
    wdGoToFootnote = 4
    wdGoToEndnote = 5
    wdGoToComment = 6
    wdGoToField = 7
    wdGoToGraphic = 8
    wdGoToObject = 9
    wdGoToEquation = 10
    wdGoToHeading = 11
    wdGoToPercent = 12
    wdGoToSpellingError = 13
    wdGoToGrammaticalError = 14
    wdGoToProofreadingError = 15


class WdHeaderFooterIndex(IntEnum):
    # https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.wdheaderfooterindex
    wdHeaderFooterPrimary = 1
    wdHeaderFooterFirstPage = 2
    wdHeaderFooterEvenPages = 3


class LineSpacingRule(IntEnum):
    # https://learn.microsoft.com/en-us/office/vba/api/word.wdlinespacing
    wdLineSpaceSingle = 0
    wdLineSpace1pt5 = 1
    wdLineSpaceDouble = 2
    wdLineSpaceAtLeast = 3
    wdLineSpaceExactly = 4
    wdLineSpaceMultiple = 5


class WdPageNumberAlignment(IntEnum):
    # https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.wdpagenumberalignment
    wdAlignPageNumberLeft = 0
    wdAlignPageNumberCenter = 1
    wdAlignPageNumberRight = 2
    wdAlignPageNumberInside = 3
    wdAlignPageNumberOutside = 4


class WdSaveOptions(IntEnum):
    # https://learn.microsoft.com/en-us/office/vba/api/word.wdsaveoptions
    wdPromptToSaveChanges = -2
    wdSaveChanges = -1
    wdDoNotSaveChanges = 0


class Document:
    """Document is a wrapper around Word.Document COM object.
    All the index used in the function is 0-based and converted to 1-based while calling COM functions.
    """

    def __init__(self, app: object, doc: object):
        self.app = app
        self.doc = doc

    def print_as(self, pathname: str, format_type=WdExportFormat.wdExportFormatPDF) -> None:
        self.doc.ExportAsFixedFormat(pathname, format_type)

    def close(self, option: WdSaveOptions = WdSaveOptions.wdDoNotSaveChanges) -> None:
        if self.doc:
            self.doc.Close(option)
            self.doc = None

    def __enter__(self) -> "Document":
        return self

    def __exit__(self, exc_type, exc_value, traceback) -> None:
        self.close()


class App:
    @staticmethod
    def is_running() -> bool:
        return process_exists("winword.exe")

    def __init__(self, visible: bool = True):
        pythoncom.CoInitialize()

        # https://stackoverflow.com/questions/50127959/win32-dispatch-vs-win32-gencache-in-python-what-are-the-pros-and-cons
        # self.word = win32com.client.gencache.EnsureDispatch("Word.Application")
        self.word = win32com.client.Dispatch("Word.Application")

        if visible:
            self.word.Visible = 1

    def new(self) -> "Document":
        doc = self.word.Documents.Add()
        if doc.Windows.Count == 0:
            doc.NewWindow()

        return Document(self, doc)

    def _find_document(self, filename: str) -> int:
        for i in range(self.word.Documents.Count):
            doc = self.word.Documents.Item(i + 1)
            if doc.Name == filename:
                return i

        return -1

    def open(self, pathname: str) -> "Document":
        index = self._find_document(os.path.split(pathname)[1])
        if index != -1:
            doc = self.word.Documents.Item(index + 1)
        else:
            self.word.DisplayAlerts = False
            doc = self.word.Documents.Open(pathname)
            self.word.DisplayAlerts = True

        return Document(self, doc)

    def quit(self, force: bool = False, only_if_empty: bool = True) -> None:
        call_quit = force
        if call_quit is False:
            if only_if_empty and self.word.Documents.Count == 0:
                call_quit = True

        if call_quit is False:
            return

        self._quit()

    def _quit(self) -> None:
        try:
            self.word.Quit()
        except AttributeError:
            pass
