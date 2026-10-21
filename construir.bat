@echo off
rem Genera dist\PPTRemoto.exe (un solo archivo para copiar a la laptop del PowerPoint)
python -m pip install -r requirements.txt pyinstaller
python -m PyInstaller --noconfirm --onefile --noconsole --name PPTRemoto --add-data "web;web" --collect-submodules uvicorn --hidden-import win32timezone ppt_remote.py
