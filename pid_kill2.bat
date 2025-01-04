@echo off
:loop
for /f "tokens=2 delims=," %%i in ('tasklist /v /fo csv ^| findstr "cmd.exe"') do taskkill /F /PID %%i
tasklist /v | findstr "cmd.exe" >nul
if not errorlevel 1 goto loop
exit