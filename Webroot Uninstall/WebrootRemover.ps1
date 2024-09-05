@echo off
REM Uninstall Webroot
"C:\Program Files (x86)\Webroot\wrsa.exe" -uninstall

REM Wait for 10 seconds
timeout /t 10 /nobreak

REM Delete Registry Keys (Not all keys are present)
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" /v WRSVC /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WRUNINST" /f
reg delete "HKLM\SOFTWARE\WRData" /f
reg delete "HKLM\SYSTEM\ControlSet001\services\WRSVC" /f
reg delete "HKLM\SYSTEM\ControlSet002\services\WRSVC" /f
reg delete "HKLM\SYSTEM\CurrentControlSet\services\WRSVC" /f

REM Delete the Service
sc.exe \\%COMPUTERNAME% delete WRSVC

REM Wait for 5 seconds
timeout /t 5 /nobreak

REM Delete the Webroot Directory
rd /s /q "C:\Program Files (x86)\Webroot"


# Variables
$webrootPath = "C:\Program Files (x86)\Webroot"
$registryPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\WRSVC",
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WRUNINST",
    "HKLM:\SOFTWARE\WRData",
    "HKLM\SYSTEM\ControlSet001\services\WRSVC",
    "HKLM\SYSTEM\ControlSet002\services\WRSVC",
    "HKLM\SYSTEM\CurrentControlSet\services\WRSVC"
)
$serviceName = "WRSVC"

# Webroot Uninstall
Start-Process "C:\Program Files (x86)\Webroot\wrsa.exe" -ArgumentList "-uninstall" -Wait

# Pause to allow for Uninstall
Start-Sleep -Seconds 10

foreach ($path in $registryPaths) {
    if (Test-Path $path) {
        Remove-Item $path -Recurse -Force
    }
}

# Remove the service from SC.exe
if (Get-Service -Name $serviceName -ErrorAction SilentlyContinue) {
    Stop-Service -Name $serviceName -Force
    Remove-Service -Name $serviceName -ErrorAction SilentlyContinue
}

Start-Sleep -Seconds 5

if (Test-Path $webrootPath) {
    Remove-Item $webrootPath -Recurse -Force
}
