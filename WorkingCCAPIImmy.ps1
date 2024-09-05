# Ensuring  TLS 1.2 is used for the connection
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Define URLs and file paths
$url = 'https://acumen.immy.bot/plugins/api/v1/1/installer/latest-download'
$InstallerFile = [io.path]::ChangeExtension([io.path]::GetTempFileName(), ".msi")
$InstallerLogFile = [io.path]::ChangeExtension([io.path]::GetTempFileName(), ".log")

# Downloading file
try {
    $webClient = New-Object System.Net.WebClient
    $webClient.DownloadFile($url, $InstallerFile)
    Write-Host "File downloaded successfully to $InstallerFile"
} catch {
    Write-Host "Error: Could not download the file."
    Write-Host "Details: $_"
    exit 1
}


# Install the file
$Arguments = "/c msiexec /i `"$InstallerFile`" /qn /norestart /l*v `"$InstallerLogFile`" REBOOT=REALLYSUPPRESS ID=IDGOESHERE@@@@@ ADDR=https://acumen.immy.bot/plugins/api/v1/1 KEY=KEYGOESHERE@@@@="
Write-Host "InstallerLogFile: $InstallerLogFile"

try {
    $Process = Start-Process -Wait cmd -ArgumentList $Arguments -Passthru
    if ($Process.ExitCode -ne 0) {
        Get-Content $InstallerLogFile -ErrorAction SilentlyContinue | Select-Object -Last 200
        throw "Exit Code: $($Process.ExitCode), ComputerName: $($env:ComputerName)"
    } else {
        Write-Host "Exit Code: $($Process.ExitCode)"
        Write-Host "ComputerName: $($env:ComputerName)"
    }
} catch {
    Write-Host "Error during installation."
    Write-Host "Details: $_"
}
