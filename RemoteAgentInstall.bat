@echo off
setlocal

:: This script will install Automate on the computer. 
:: It will pull the Automate Installer .MSI and then transform the configuration with the AutomateMST.MST file from a cloud-hosted url (Azure Blobs, etc.). 
:: It must be ran as an Administrator.
:: Written by: Jesse Campanella

:: Define URLs and local paths
set "msiUrl=http://testurl.com/AutomateAgent.MSI"
set "mstUrl=http://testurl.com/AutomateMST.MST"
set "msiLocalPath=C:\temp\AutomateAgent.MSI"
set "mstLocalPath=C:\temp\AutomateMST.MSI"

:: Create the temp directory if it doesn't exist
if not exist "C:\temp" mkdir C:\temp

:: Download the .msi file
powershell -Command "Invoke-WebRequest -Uri '%msiUrl%' -OutFile '%msiLocalPath%'"

:: Download the .mst file
powershell -Command "Invoke-WebRequest -Uri '%mstUrl%' -OutFile '%mstLocalPath%'"

:: Execute the install
msiexec /i '%msiLocalPath%' TRANSFORMS='%mstLocalPath%' /quiet /norestart

endlocal