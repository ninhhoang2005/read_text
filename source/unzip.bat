@echo off
cd /d "%~dp0"

title Updating is in progress, please do not close this window

if not exist "read_text.zip" (
    echo Error: File read_text.zip not found!
    pause
    exit
)

powershell -command "Expand-Archive -Path 'read_text.zip' -DestinationPath '.' -Force"

if exist "read_text.zip" del "read_text.zip"

start "" "ReadText.exe"

exit
