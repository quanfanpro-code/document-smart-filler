@echo off
chcp 65001 >nul
set "FILLER_PYTHON=%LOCALAPPDATA%\Programs\Python\Python314\python.exe"
if not exist "%FILLER_PYTHON%" (
 echo 没有找到 Python 3.14。请先按照 readme.md 的“第一次使用”说明安装。
 pause
 exit /b 1
)
cd /d "%~dp0"
"%FILLER_PYTHON%" -c "import customtkinter,docx,openpyxl,requests,fitz,PIL,win32com" >nul 2>nul
if errorlevel 1 (
 echo 缺少运行组件。请双击“安装依赖.bat”，完成后再打开本程序。
 pause
 exit /b 1
)
start "" "%LOCALAPPDATA%\Programs\Python\Python314\pythonw.exe" "%~dp0启动.py"