import logging
import os
import sys
from typing import Optional


class AnsiTermFormatter(logging.Formatter):

    BLACK = "\x1b[30m"
    RED = "\x1b[31m"
    GREEN = "\x1b[32m"
    YELLOW = "\x1b[33m"
    BLUE = "\x1b[34m"
    MAGENTA = "\x1b[35m"
    CYAN = "\x1b[36m"
    WHITE = "\x1b[37m"

    BRIGHT_BLACK = "\x1b[90m"
    BRIGHT_RED = "\x1b[91m"
    BRIGHT_GREEN = "\x1b[92m"
    BRIGHT_YELLOW = "\x1b[93m"
    BRIGHT_BLUE = "\x1b[94m"
    BRIGHT_MAGENTA = "\x1b[95m"
    BRIGHT_CYAN = "\x1b[96m"
    BRIGHT_WHITE = "\x1b[97m"

    RESET = "\x1b[0m"
    COLOR = {
        logging.DEBUG: WHITE,
        logging.INFO: BRIGHT_WHITE,
        logging.WARNING: BRIGHT_YELLOW,
        logging.ERROR: BRIGHT_RED,
        logging.CRITICAL: BRIGHT_MAGENTA,
    }

    @staticmethod
    def is_ansi_color_term() -> bool:
        # Need to check: # https://learn.microsoft.com/en-us/windows/console/setconsolemode
        # but turn it on for now.
        if sys.platform.startswith("win32"):
            return True
        else:
            return os.environ.get("TERM") in ["xterm-256color", "xterm-16color", "xterm-color"]

    def __init__(self, formatter: Optional[logging.Formatter] = None):
        if formatter:
            self._style = formatter._style
            self._fmt = formatter._fmt
            self.datefmt = formatter.datefmt

        if AnsiTermFormatter.is_ansi_color_term():
            self.formatMessage = self._formatMessage  # type: ignore

    def _formatMessage(self, record: logging.LogRecord) -> str:
        message = super(AnsiTermFormatter, self).formatMessage(record)
        log_fmt = AnsiTermFormatter.COLOR.get(record.levelno)
        if log_fmt:
            message = log_fmt + message + AnsiTermFormatter.RESET
        return message


def init_log(verbose: bool = False) -> None:
    log_level = logging.DEBUG if verbose else logging.INFO

    logging.basicConfig(stream=sys.stdout, format="%(asctime)s - %(levelname)s - %(message)s")
    root = logging.getLogger()
    root.setLevel(log_level)
    if AnsiTermFormatter.is_ansi_color_term():
        root.handlers[0].setFormatter(AnsiTermFormatter(root.handlers[0].formatter))
    logging.Formatter.default_msec_format = "%s.%03d"
