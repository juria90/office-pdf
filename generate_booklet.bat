@echo off
:: skip "Active code page: "
FOR /F "tokens=4" %%g IN ('chcp') do SET PREVIOUS_CHCP=%%g
:: echo codepage=%PREVIOUS_CHCP%

:: Do not use CHCP that clears the screen, but use mode.com: https://github.com/microsoft/terminal/issues/9446
mode.com con cp select=65001 > NUL

setlocal

set PROG_DIR=%HOMEDRIVE%%HOMEPATH%\Church\office-pdf

cd /d %PROG_DIR%

if not exist venv\Scripts\activate.bat (
    echo Installing venv and requirements.
    python3 -m venv venv

    echo Activating venv.
	call venv\Scripts\activate.bat
	pip install -r requirements.txt
) else (
    echo Activating venv.
	call venv\Scripts\activate.bat
)

@echo on
REM python emc_booklet.py --mdb-filename C:\Church\Address\address.mdb --booklet-filename C:\Users\jameslee-pc\Dropbox\EMC\일반행정\2024\2024 신앙생활요람.docx
python emc_booklet.py --mdb-filename "C:\Users\jameslee-pc\Church\office-pdf\address.mdb" --booklet-filename "C:\Users\jameslee-pc\Dropbox\EMC\일반행정\2024\2024 신앙생활요람.docx"

pause

endlocal

@echo off
mode.com con cp select=%PREVIOUS_CHCP% > NUL
