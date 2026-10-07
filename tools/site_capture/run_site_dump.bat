@echo off
python "%~dp0..\..\scripts\capture_site.py" %*
exit /b %errorlevel%
