@echo off
title "Updating is in progress, please do not close this window"
powershell -command "Expand-Archive -Path 'read_text.zip' -DestinationPath '.' -Force"
del read_text.zip
start "" read_text.exe
exit
