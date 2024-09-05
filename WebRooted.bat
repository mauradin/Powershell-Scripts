@echo off
REM Uninstall Webroot
"C:\Program Files (x86)\Webroot\wrsa.exe" -uninstall

REM Wait for 30 seconds
timeout /t 30 /nobreak

REM Delete Registry Keys (Not all keys are present)
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" /v WRSVC /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WRUNINST" /f
reg delete "HKLM\SOFTWARE\WRData" /f
reg delete "HKLM\SYSTEM\ControlSet001\services\WRSVC" /f
reg delete "HKLM\SYSTEM\ControlSet002\services\WRSVC" /f
reg delete "HKLM\SYSTEM\CurrentControlSet\services\WRSVC" /f

REM Delete the Service
sc.exe \\%COMPUTERNAME% delete WRSVC

REM Wait for 3 seconds
timeout /t 5 /nobreak

REM Delete the Webroot Directory
rd /s /q "C:\Program Files (x86)\Webroot"
rd /s /q "C:\ProgramData\WRCore"
rd /s /q "C:\ProgramData\WRData"