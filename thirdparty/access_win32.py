from enum import IntEnum

import pythoncom  # type: ignore
import win32com.client  # type: ignore

from process_exists import process_exists


pythoncom.CoInitialize()

# https://learn.microsoft.com/en-us/office/vba/api/overview/access


class AcFormat(IntEnum):
    acFormatASP = 0
    acFormatHTML = 1
    acFormatIIS = 2
    acFormatPDF = 3
    acFormatRTF = 4
    acFormatSNP = 5
    acFormatTXT = 6
    acFormatXLS = 7
    acFormatXPS = 8


# https://learn.microsoft.com/en-us/previous-versions/office/office-12//bb226001(v=office.12)
# https://answers.microsoft.com/en-us/msoffice/forum/all/variable-not-defined-on-outputto-acformatpdf/7d6f9cdd-6ca5-4716-b688-bdd6f3cea831
AcFormatMap = {
    AcFormat.acFormatASP: "Microsoft",  # tables, queries, and forms
    AcFormat.acFormatHTML: "HTML",
    AcFormat.acFormatIIS: "Microsoft",  # tables, queries, and forms
    AcFormat.acFormatPDF: "PDF",  # "PDF Format (*.pdf)" works
    AcFormat.acFormatRTF: "Rich",
    AcFormat.acFormatSNP: "Snapshot",
    AcFormat.acFormatTXT: "MS-DOS",
    AcFormat.acFormatXLS: "Microsoft",
    AcFormat.acFormatXPS: "XPS",  # "XPS Format (*.xps)" is not verified.
}


# https://learn.microsoft.com/en-us/office/vba/api/access.acoutputobjecttype
class AcOutputObjectType(IntEnum):
    acOutputTable = 0
    acOutputQuery = 1
    acOutputForm = 2
    acOutputReport = 3
    acOutputModule = 5
    acOutputServerView = 7
    acOutputStoredProcedure = 9
    acOutputFunction = 10


# https://learn.microsoft.com/en-us/office/vba/api/access.acquitoption
class AcQuitOption(IntEnum):
    acQuitPrompt = 0
    acQuitSaveAll = 1
    acQuitSaveNone = 2


# https://learn.microsoft.com/en-us/office/vba/api/access.acview
class AcView(IntEnum):
    acViewNormal = 0
    acViewDesign = 1
    acViewPreview = 2
    acViewPivotTable = 3
    acViewPivotChart = 4
    acViewReport = 5
    acViewLayout = 6


class AccessObj:
    def __init__(self, app: "App", name: str):
        super().__setattr__("app", app)
        super().__setattr__("name", name)
        super().__setattr__("obj", None)

        # self.app = app
        # self.name = name
        # self.obj = None

    def close(self) -> None:
        if self.name:
            self.name = ""
            self.obj = None

    def __getattr__(self, attr):
        sup = super()
        if hasattr(sup, attr):
            return sup.__getattr__(attr)
        elif (obj := sup.__getattr__("obj")) is not None:
            return getattr(obj, attr)

    def __setattr__(self, attr, value):
        sup = super()
        if hasattr(self, attr):
            return sup.__setattr__(attr, value)
        elif (obj := self.obj) is not None:
            return setattr(obj, attr, value)

    def __enter__(self) -> "AccessObj":
        return self

    def __exit__(self, exc_type, exc_value, traceback) -> None:
        self.close()


class Report(AccessObj):
    """Report is a wrapper around Access.Report COM object.
    All the index used in the function is 0-based and converted to 1-based while calling COM functions.
    """

    def __init__(self, app: "App", name: str):
        super().__init__(app, name)

        reports = self.app.access.Reports
        for i in range(reports.Count):
            # zero based index: https://learn.microsoft.com/en-us/office/vba/api/access.reports
            report = reports.Item(i)
            if report.Name == self.name:
                self.obj = report
                break

    def print_as(self, pathname: str, format_type=AcFormat.acFormatPDF) -> None:
        self.app.print_report_as(self.name, pathname, format_type)


class Table(AccessObj):
    """Table is a wrapper around Access.Table COM object.
    All the index used in the function is 0-based and converted to 1-based while calling COM functions.
    """

    def __init__(self, app: "App", name: str):
        super().__init__(app, name)

    def print_as(self, pathname: str, format_type=AcFormat.acFormatPDF) -> None:
        self.app.print_table_as(self.name, pathname, format_type)


class App:
    @staticmethod
    def is_running() -> bool:
        return process_exists("access.exe")

    def __init__(self, visible: bool = True) -> None:
        # https://stackoverflow.com/questions/50127959/win32-dispatch-vs-win32-gencache-in-python-what-are-the-pros-and-cons
        # self.access = win32com.client.gencache.EnsureDispatch("Access.Application")
        self.access = win32com.client.Dispatch("Access.Application")

        if visible:
            self.access.Visible = 1

    def open(self, pathname: str) -> bool:
        # curProj = self.access.Currentproject.path  # returns dir
        curDB = self.access.CurrentDb()
        cur_name = curDB.name if curDB is not None else ""
        if cur_name != pathname:
            if cur_name:
                self.access.CloseCurrentDatabase()
            self.access.OpenCurrentDatabase(pathname)

        return True

    def open_report(self, report_name: str, view: AcView = AcView.acViewDesign) -> "Report":
        self.access.DoCmd.OpenReport(report_name, view)
        return Report(self, report_name)

    def open_table(self, table_name: str, view: AcView = AcView.acViewNormal) -> "Table":
        self.access.DoCmd.OpenTable(table_name, view)
        return Table(self, table_name)

    def print_report_as(self, report_name: str, pathname: str, format_type=AcFormat.acFormatPDF) -> None:
        self.access.DoCmd.OutputTo(AcOutputObjectType.acOutputReport, report_name, AcFormatMap[format_type], pathname)

    def print_table_as(self, table_name: str, pathname: str, format_type=AcFormat.acFormatPDF) -> None:
        self.access.DoCmd.OutputTo(AcOutputObjectType.acOutputTable, table_name, AcFormatMap[format_type], pathname)

    def quit(self, option: AcQuitOption = AcQuitOption.acQuitSaveAll) -> None:
        try:
            self.access.Quit(option)
        except AttributeError:
            pass
