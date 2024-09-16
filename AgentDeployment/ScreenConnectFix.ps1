# If a device is in Automate, and not ScreenConnect - this script will utilize remote Powershell in Automate to install ScreenConnect

$url = "URL.MSI"
$localPath = "$env:TEMP\ScreenConnect.ClientSetup.msi"
Write-Output "Downloading file from $url..."
Invoke-WebRequest -Uri $url -OutFile $localPath
if (Test-Path $localPath) {
    Write-Output "File downloaded successfully to $localPath."
    Write-Output "Starting the installation..."
    Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$localPath`" /quiet /norestart" -Wait
    Remove-Item -Path $localPath -Force

    Write-Output "Installation complete."
} else {
    Write-Output "Failed to download the file."
}
