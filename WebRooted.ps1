$repositoryUrl = "https://atworktest.blob.core.windows.net/aworkcsdeployment/WebRooted.bat"
$destinationPath = "C:\webrooted\WebRooted.bat"

# Create destination
if (-not (Test-Path -Path (Split-Path -Path $destinationPath -Parent))) {
    Write-Output "Destination directory does not exist. Creating..."
    New-Item -Path (Split-Path -Path $destinationPath -Parent) -ItemType Directory -Force
}

# Download WebRooted.Bat
$ProgressPreference = 'SilentlyContinue'
Invoke-WebRequest -Uri $repositoryUrl -OutFile $destinationPath

# Log for file download
if (Test-Path $destinationPath) {
    Write-Output "File downloaded successfully. Installing..."

    # Run the .bat
    Start-Process -FilePath $destinationPath -Wait -NoNewWindow

    Write-Output "Installation complete."
} else {
    Write-Output "File download failed."
}