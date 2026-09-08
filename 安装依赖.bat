@echo off
chcp 65001 >nul
set "FILLER_PYTHON=%LOCALAPPDATA%\Programs\Python\Python314\python.exe"
if not exist "%FILLER_PYTHON%" (
 echo 请先安装 Python 3.14，并使用安装程序默认的当前用户安装位置。
 pause
 exit /b 1
)
"%FILLER_PYTHON%" -m pip install -r "%~dp0requirements.txt"
if errorlevel 1 (
 echo 组件安装未完成，请检查网络，并保留上面的错误信息。
) else (
 echo 安装完成，现在可以双击“启动程序.bat”。
)
pause