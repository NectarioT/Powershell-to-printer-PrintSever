@echo off
REM Wrapper that runs Connect-ADPrintServer.ps1 with the local execution
REM policy bypassed for this single invocation. All arguments are forwarded
REM to the script.
REM
REM Examples:
REM   Run-Connect-ADPrintServer.cmd -Test
REM   Run-Connect-ADPrintServer.cmd -Test -LocalIP 10.26.26.47
REM   Run-Connect-ADPrintServer.cmd -PrintServer PRINTSRV01 -AutoDetect

setlocal
set "SCRIPT_DIR=%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%Connect-ADPrintServer.ps1" %*
set "EXITCODE=%ERRORLEVEL%"
endlocal & exit /b %EXITCODE%
