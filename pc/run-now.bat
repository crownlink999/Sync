@echo off
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0svodki-collect.ps1"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0svodki-send.ps1"
pause
