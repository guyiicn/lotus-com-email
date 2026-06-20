@echo off
REM Capture the current foreground window. PowerShell cannot bring a window
REM to front reliably from a detached cmd, so we rely on the window that the
REM user already has focused (the ZCode window they are typing into).
setlocal
set OUT=%~dp0attachments\window_%random%%random%.png
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0screenshot.ps1" -OutFile "%OUT%"
echo %OUT%
endlocal
