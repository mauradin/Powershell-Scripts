$FontsDir = "C:\Program Files\Adobe\Acrobat DC\Resource\Font"
$ZipFilePath = "C:\Program Files\Adobe\Acrobat DC\Resource\Font\ETCOMFonts.zip"
$DownloadUrl = "OneDriveURL?download=1 "

if (-not (Test-Path $FontsDir)) {
    New-Item -Path $FontsDir -ItemType Directory
} else {
}

if (Test-Path $ZipFilePath) {
} else {
    
    try {
        Invoke-WebRequest -Uri $DownloadUrl -OutFile $ZipFilePath -TimeoutSec 300
    } catch {
        exit 1
    }
}

Expand-Archive -Path $ZipFilePath -DestinationPath $FontsDir -Force
Start-Sleep -Seconds 10
if (Test-Path $ZipFilePath) {
    Remove-Item -Path $ZipFilePath -Force
} else {
}