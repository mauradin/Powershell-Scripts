$fortiPath = "C:\Program Files\Fortinet\FortiClient"
$xmlFilePath = Join-Path -Path $fortiPath -ChildPath "VPNConfig.xml"

$xmlContent = @"
<ADD .XML File HERE>
"@

$xmlContent = $xmlContent -replace "`n", "`r`n"

$utf8NoBomEncoding = New-Object System.Text.UTF8Encoding $False
[System.IO.File]::WriteAllText($xmlFilePath, $xmlContent, $utf8NoBomEncoding)

Start-Sleep -Seconds 2

$cmdString = "cd `"$fortiPath`" && fcconfig.exe -f VPNConfig.xml -m all -o import -p 123456789"
Start-Process -Verb RunAs -FilePath "cmd.exe" -ArgumentList "/k $cmdString"