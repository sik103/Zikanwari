@echo off

set PATH=\WinPython-64bit-3.5.4.1Qt5\python-3.5.4.amd64;%PATH%
set /P STR_INPUT="zikanwari2.py (1) or z4jpg.py (0) [1]: "

IF "%STR_INPUT%" == "0" (
    python z4jpg.py
) ELSE (
    python zikanwari2.py
)

pause
