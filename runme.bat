@echo off
:: Change directory to the specific folder
cd "%~dp0"

:: Hide the Python file to make it invisible
attrib +h runthis.py

:: Run the Python script using Python
py runthis.py

:: Pause to keep the window open after execution
pause
