REM ***************************************************************************************
REM Autor:		        mdo
REM Version:	        1.2
REM Usage:		        This script shall Copy the Setup Script with Resources into c:\root\
REM TODO:               Config nächstes gerät auf x setzen
REM ***************************************************************************************
PowerShell -NoProfile -ExecutionPolicy Bypass -Command "& {Start-Process PowerShell -ArgumentList 'Set-ExecutionPolicy Unrestricted -Force' -Verb RunAs}"

mkdir "c:\root\"
mkdir "c:\root\Setup\"

xcopy %~dp0 "c:\root\Setup" /E /C /R
del c:\root\Setup\init.bat
pause