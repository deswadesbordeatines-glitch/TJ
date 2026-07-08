@echo off
setlocal

set APP_NAME=发票识别工具
set PROJECT_DIR=%~dp0
cd /d "%PROJECT_DIR%"

where python >nul 2>nul
if errorlevel 1 (
  echo 未找到 Python。请在打包用的 Windows 电脑上安装 Python 3.10+。
  exit /b 1
)

python -c "import tkinter"
if errorlevel 1 (
  echo 当前 Python 缺少 tkinter，无法打包桌面程序。
  exit /b 1
)

if exist .build-venv rmdir /s /q .build-venv
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

python -m venv .build-venv
call .build-venv\Scripts\activate.bat
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

pyinstaller ^
  --noconfirm ^
  --clean ^
  --onedir ^
  --windowed ^
  --name "%APP_NAME%" ^
  main.py

copy README.txt "dist\%APP_NAME%\README.txt" >nul

echo.
echo 打包完成：
echo %PROJECT_DIR%dist\%APP_NAME%\%APP_NAME%.exe
echo.
echo 请把整个文件夹复制到离线电脑：
echo %PROJECT_DIR%dist\%APP_NAME%

endlocal
