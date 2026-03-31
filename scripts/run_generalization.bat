@echo off
setlocal

REM Usage:
REM   scripts\run_generalization.bat [SAMPLE_DIR] [MAX_UNKNOWN_RATIO] [REPORT_PATH]
REM Example:
REM   scripts\run_generalization.bat E:\docx-formatter\samples\generalization 0.03 memory-bank\baseline\generalization_summary.json

set "SAMPLE_DIR=%~1"
if "%SAMPLE_DIR%"=="" set "SAMPLE_DIR=%CD%\samples\generalization"

set "MAX_UNKNOWN_RATIO=%~2"
if "%MAX_UNKNOWN_RATIO%"=="" set "MAX_UNKNOWN_RATIO=0.03"

set "REPORT_PATH=%~3"
if "%REPORT_PATH%"=="" set "REPORT_PATH=memory-bank\baseline\generalization_summary.json"

set "DOCX_GENERALIZATION_DIR=%SAMPLE_DIR%"
set "DOCX_GENERALIZATION_MAX_UNKNOWN_RATIO=%MAX_UNKNOWN_RATIO%"
set "DOCX_GENERALIZATION_REPORT_PATH=%REPORT_PATH%"

echo [generalization] sample_dir=%DOCX_GENERALIZATION_DIR%
echo [generalization] max_unknown_ratio=%DOCX_GENERALIZATION_MAX_UNKNOWN_RATIO%
echo [generalization] report_path=%DOCX_GENERALIZATION_REPORT_PATH%

python -m unittest discover -s tests -p test_generalization_suite.py -v
set "EXIT_CODE=%ERRORLEVEL%"

if %EXIT_CODE% NEQ 0 (
  echo [generalization] FAILED with exit code %EXIT_CODE%
  endlocal & exit /b %EXIT_CODE%
)

echo [generalization] DONE
endlocal & exit /b 0
