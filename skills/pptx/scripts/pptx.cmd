@echo off
where py >nul 2>nul
if %errorlevel%==0 goto run_py
where python >nul 2>nul
if %errorlevel%==0 goto run_python
echo pptx requires Python 3.10 or newer. 1>&2
exit /b 1

:run_py
py -3 -X utf8 "%~dp0pptx_cli.py" %*
exit /b %errorlevel%

:run_python
python -X utf8 "%~dp0pptx_cli.py" %*
exit /b %errorlevel%
