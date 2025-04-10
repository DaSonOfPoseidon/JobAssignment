@echo off
cd /d %~dp0

echo Setting up local environment...
python-embed\python.exe -m ensurepip --default-pip
python-embed\python.exe -m pip install --upgrade pip

echo Installing dependencies...
python-embed\python.exe -m pip install -r requirements.txt --target=python-embed\lib

echo Launching the tool...
python-embed\python.exe ASSigner.py

pause