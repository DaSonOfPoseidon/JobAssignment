@echo off
cd /d %~dp0

REM === SETUP LOCAL PYTHON ENVIRONMENT ===
IF NOT EXIST "embedded_python\Scripts\pip.exe" (
    echo Setting up pip in embedded Python...
    embedded_python\python.exe -m ensurepip --default-pip
)

REM === INSTALL MISSING DEPENDENCIES ONLY ===
echo Checking Python packages...
SETLOCAL ENABLEDELAYEDEXPANSION
SET PACKAGES=pandas selenium openpyxl webdriver-manager tkinterdnd2
FOR %%P IN (%PACKAGES%) DO (
    embedded_python\python.exe -c "import %%P" 2>NUL
    IF ERRORLEVEL 1 (
        echo Installing: %%P
        embedded_python\python.exe -m pip install %%P --target=embedded_python\lib
    ) ELSE (
        echo Found: %%P
    )
)
ENDLOCAL

REM === LAUNCH THE TOOL ===
echo Starting tool...
embedded_python\python.exe assign_workorders_gui.py

pause
