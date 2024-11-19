Stop-Process -Name "outlook" -Force -ErrorAction SilentlyContinue

Get-ChildItem -Path "C:\Users" -Directory | ForEach-Object {
    $userDir = $_.FullName

    $outlookFolder = Join-Path -Path $userDir -ChildPath "AppData\Local\Microsoft\Outlook"
    if (Test-Path -Path $outlookFolder) {
        Rename-Item -Path $outlookFolder -NewName "Outlook.oldmigration" -Force
        Write-Host "Renamed folder for user $userDir"
    }

    $oneAuthFolder = Join-Path -Path $userDir -ChildPath "AppData\Local\Microsoft\OneAuth"
    if (Test-Path -Path $oneAuthFolder) {
        Rename-Item -Path $oneAuthFolder -NewName "OneAuth.oldmigration" -Force
        Write-Host "Renamed folder for user $userDir"
    }
}

$baseRegistryPath = "Software\Microsoft\Office\16.0\Outlook\Profiles"

$hkeyUsers = Get-ChildItem -Path "Registry::HKEY_USERS"

foreach ($user in $hkeyUsers) {
    $userRegistryPath = Join-Path -Path $user.PSPath -ChildPath $baseRegistryPath

    if (Test-Path $userRegistryPath) {
        Remove-Item -Path $userRegistryPath -Recurse -Force
        Write-Host "Deleted Outlook Profiles registry key for user: $($user.PSChildName)"
    } else {
        Write-Host "No Outlook Profiles registry key found for user: $($user.PSChildName)"
    }
}
