@echo off
REM Launch the lotus-notes-mail MCP server.
REM Uses the bundled 32-bit Python (required for Lotus Notes COM).
REM ZCode / Claude Desktop / Cursor call this as a stdio MCP server.

setlocal
set ROOT=%~dp0
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8
"%ROOT%python\python.exe" "%ROOT%src\server.py"
endlocal
