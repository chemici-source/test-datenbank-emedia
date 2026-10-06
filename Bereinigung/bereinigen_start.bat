@echo off
REM Doppelklick reicht nicht - diese Datei zieht man eine Excel-Datei drauf
REM (Datei aus dem Explorer per Drag-and-Drop auf dieses Icon ziehen).
REM Voraussetzung: Python ist installiert und "bereinigen.py" liegt im
REM gleichen Ordner wie diese .bat-Datei.

if "%~1"=="" (
    echo Bitte eine Excel-Datei auf dieses Icon ziehen und fallen lassen.
    pause
    exit /b
)

python "%~dp0bereinigen.py" "%~1"
pause
