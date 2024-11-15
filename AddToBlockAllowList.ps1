#This script will pipe domains from a text file into the Allow/Block list for Mail Flow in Microsoft Defender
Import-Module ExchangeOnline
Connect-ExchangeOnline


$domainPath = "C:\temp\domains.txt"
$domains = Get-Content -Path $domainPath

if ($domains.Count -eq 0) {
    Write-Host "Not working for some reason :)."
    exit
}

foreach ($domain in $domains) {
    $domain = $domain.Trim()
# Removing whitespace/trailing
    if ([string]::IsNullOrWhiteSpace($domain)) {
        continue
    }

    try {
        Write-Host "Thomas is adding this domain: $domain"
        New-TenantAllowBlockListItems -ListType Sender -Block -Entries $domain -NoExpiration
        Write-Host "Successfully added $domain to the Block list."
    } catch {
        Write-Host "Failed to add $domain. Error: $_"
    }
}
Disconnect-ExchangeOnline -Confirm:$false
