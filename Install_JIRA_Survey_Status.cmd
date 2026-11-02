@echo off
setlocal
cd /d "%~dp0"
REM This sets execution policy ONLY for the PowerShell process started by this installer.
REM It does not change user/machine execution policy. Organization policies still apply.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Build_and_Install_JIRA_Survey_Status.ps1"
set "INSTALL_EXIT=%ERRORLEVEL%"
echo.
if not "%INSTALL_EXIT%"=="0" (
    echo INSTALLATION FAILED. See the error details above.
    echo If your organization enforces script restrictions through Group Policy, contact IT.
) else (
    echo Installation command completed. Review any Excel messages above.
)
pause
exit /b %INSTALL_EXIT%
