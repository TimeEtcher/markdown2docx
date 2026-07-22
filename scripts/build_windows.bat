@echo off
REM md2word Windows 一键打包脚本
REM 前提：已安装 Python 3.10+，并在本目录执行
cd /d "%~dp0\.."

python -m venv .venv
call .venv\Scripts\activate.bat
pip install -r requirements.txt

pyinstaller --onefile --windowed --name md2word --collect-data latex2mathml run.py

echo.
echo 打包完成: dist\md2word.exe
pause
