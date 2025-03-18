@echo off
set PYTHON_DIR=%USERPROFILE%\Downloads\python-3.10.8-amd64-portable
set PATH=%PYTHON_DIR%;%PYTHON_DIR%\Scripts

python fetch_vcdist.py --destination=%cd% --include-old-versions
pause
