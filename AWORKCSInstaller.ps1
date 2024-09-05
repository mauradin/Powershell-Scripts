$repositoryUrl = "https://atworktest.blob.core.windows.net/aworkcsdeployment/Crowdstrike.exe"
$destinationPath = "C:\crowdstrikeinstall\Crowdstrike.exe"
$cid = "CIDGOESHERE@@@@"

# Download the file
$ProgressPreference = 'SilentlyContinue'
Invoke-WebRequest -Uri $repositoryUrl -OutFile $destinationPath

# Check if the file was downloaded successfully
if (Test-Path $destinationPath) {
    Write-Output "File downloaded successfully. Installing..."

    # Install the downloaded file
    Start-Process -FilePath $destinationPath -ArgumentList "/quiet /norestart CID=$cid" -Wait -NoNewWindow

    Write-Output "Installation complete."
} else {
    Write-Output "File download failed. Please check the URL and try again."
}