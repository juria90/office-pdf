import errno
import os
import tempfile
from typing import List, Union


def InchesToPoint(i: float) -> int:
    return int(i * 72)


class BasePDFCmd:
    """Base class for PDF command classes."""

    def __init__(self, output_files: Union[str, List[str]] = [], autodelete=False) -> None:
        self.output_files: Union[str, List[str]] = output_files  # can output multiple files.
        self.autodelete = autodelete

    def __del__(self):
        self.close()

    def close(self):
        if self.autodelete:
            if isinstance(self.output_files, list):
                for fn in self.output_files:
                    self.remove_safely(fn)
            elif isinstance(self.output_files, str) and self.output_files:
                self.remove_safely(self.output_files)

    def execute(self):
        """Execute the command and produce the self.output_files files."""
        self.create_output_filename()

        self._execute()

    def _execute(self) -> None:
        pass

    def create_output_filename(self):
        create_a_new_file = False
        if isinstance(self.output_files, list):
            if len(self.output_files) == 0:
                create_a_new_file = True
        elif isinstance(self.output_files, str) and not self.output_files:
            create_a_new_file = True

        if create_a_new_file:
            ntf = tempfile.NamedTemporaryFile(mode="w+b", suffix=".pdf")
            ntf.close()

            self.output_files = ntf.name
            self.autodelete = True

        return self.output_files

    @staticmethod
    def remove_safely(filename: str):
        try:
            os.remove(filename)
        except OSError as e:
            if e.errno != errno.ENOENT:  # errno.ENOENT = no such file or directory
                raise  # re-raise exception if a different error occurred
