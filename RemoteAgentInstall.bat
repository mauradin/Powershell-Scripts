@echo off
setlocal

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