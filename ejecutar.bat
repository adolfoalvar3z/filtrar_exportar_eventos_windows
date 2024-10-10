@echo off
python -m venv .venv
CALL .venv\Scripts\activate.bat
py -m pip install --upgrade pip
pip install -r requerimientos.txt
cls
@echo on
py eventos.py