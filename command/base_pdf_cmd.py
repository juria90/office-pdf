import errno
import os
import tempfile
from typing import List


def InchesToPoint(i: float) -> int:
    return int(i * 72)


class BasePDFCmd:
    """Base class for PDF command classes."""

    def __init__(self, output_file: str = "", autodelete=False) -> None:
        self.output_file: str = ""
        self.output_file: str = output_file
        self.autodelete = autodelete

    def __del__(self):
        self.close()

    def close(self):
        if self.autodelete:
            if self.output_file:
                self.remove_safely(self.output_file)

    def execute(self):
        """Execute the command and produce the self.output_files files."""
        self.create_output_filename()

        self._execute()

    def _execute(self) -> None:
        pass

    def create_output_filename(self):
        create_a_new_file = False
        if not self.output_file:
            create_a_new_file = True

        if create_a_new_file:
            ntf = tempfile.NamedTemporaryFile(mode="w+b", suffix=".pdf")
            ntf.close()

            self.output_file = ntf.name
            self.autodelete = True

        return self.output_file

    @staticmethod
    def remove_safely(filename: str):
        try:
            os.remove(filename)
        except OSError as e:
            if e.errno != errno.ENOENT:  # errno.ENOENT = no such file or directory
                raise  # re-raise exception if a different error occurred


class BaseMultiPDFCmd(BasePDFCmd):
    """Base class for PDF command class with multiple output."""

    def __init__(self, output_files: List[str] = [], autodelete=False) -> None:
        super().__init__()

        self.output_files: List[str] = output_files
        self.autodelete = autodelete

    def close(self):
        if self.autodelete:
            for fn in self.output_files:
                self.remove_safely(fn)
