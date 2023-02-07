import logging
import os
from typing import List

from command.base_pdf_cmd import BasePDFCmd
from thirdparty.access_win32 import App as AccessApp


logger = logging.getLogger(__name__)


class PrintoutConfig:
    def __init__(self, report: str, query: str, order_by: str, output_filename: str):
        self.report = report
        self.query = query
        self.order_by = order_by
        self.output_filename = output_filename


class AccessPDFCmd(BasePDFCmd):
    def __init__(self, mdb_filename: str, printout_configs: List[PrintoutConfig]):
        super().__init__()

        self.mdb_filename = mdb_filename
        self.printout_configs = printout_configs

    def _execute(self):
        logger.info(f"Generating report files from '{self.mdb_filename}'.")

        app = AccessApp(False)
        app.open(self.mdb_filename)

        # Set papersize to statement in printer settings

        saved_query = ""
        saved_order_by = ""
        output_files = []
        for i, print_config in enumerate(self.printout_configs):
            # After printing, the report object is no longer valid design. So open and close for each print.
            # Change "Record Source", "Order By".
            with app.open_report(print_config.report) as report:
                if i == 0:
                    saved_query = report.obj.RecordSource
                    saved_order_by = report.obj.OrderBy
                report.RecordSource = print_config.query
                report.obj.OrderBy = print_config.order_by
                pdf_filename = print_config.output_filename
                try:
                    os.remove(pdf_filename)
                except OSError:
                    pass
                report.print_as(pdf_filename)

                output_files.append(pdf_filename)

        if len(self.printout_configs) > 0:
            print_config = self.printout_configs[-1]
            with app.open_report(print_config.report) as report:
                report.obj.RecordSource = saved_query
                report.obj.OrderBy = saved_order_by

        app.quit()

        self.output_files = output_files
        self.autodelete = False
