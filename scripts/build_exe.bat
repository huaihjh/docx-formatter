@echo off
setlocal

set "ROOT_DIR=%~dp0.."
pushd "%ROOT_DIR%"

if not exist ".venv\Scripts\python.exe" (
    echo Missing virtual environment: .venv\Scripts\python.exe
    exit /b 1
)

call ".venv\Scripts\python.exe" -m PyInstaller --noconfirm --clean "docx_formatter.spec"
set "EXIT_CODE=%ERRORLEVEL%"

popd
exit /b %EXIT_CODE%
