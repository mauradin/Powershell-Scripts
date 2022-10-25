<#



#>

# Gather Variables / Setup
Start-Transcript -Path "C:\AVPNOutput\transcript.txt"
Write-Host "This script will generate a new, uniquely name, root certificate and child certificate to be used for renewing and installing Azure Virtual Network Gateways." -ForegroundColor Cyan
Start-Sleep 1
Write-Host "Run this as administrator to ensure no issues." -ForegroundColor Cyan
Start-Sleep 1
$nickname = Read-Host "Enter the nickname for the cert. Format = P2SCert-'Nickname'" 
Start-Sleep 1
$expiration = Read-Host "Enter the expiration, in months, for the cert"
Start-Sleep 1
$certpasswordprompt = Read-Host "Enter the private key for installing the Client-Side VPN Child Cert."
$securepassword = ConvertTo-SecureString -String "$certpasswordprompt" -Force -AsPlainText
Start-Sleep 1
$path = "AzureVPNExport_$nickname"
Write-Host "Certs will be generated for a duration of $expiration months, under the name of P2SRootCert-$nickname, and P2SChildCert-$nickname. 
They will be exported to $ENV:USERPROFILE\Documents\$path\"
Read-Host "Press Enter to Continue"



# Generate Root Cert

$cert = New-SelfSignedCertificate -Type Custom -KeySpec Signature `
-Subject "CN=P2SRootCert-$nickname" -KeyExportPolicy Exportable `
-HashAlgorithm sha256 -KeyLength 2048 `
-NotAfter (Get-Date).AddMonths(39)`
-CertStoreLocation "Cert:\CurrentUser\My" -KeyUsageProperty Sign -KeyUsage CertSign

# Generate Child Cert
$cert2 = New-SelfSignedCertificate -Type Custom -DnsName P2SChildCert -KeySpec Signature `
-Subject "CN=P2SChildCert-$nickname" -KeyExportPolicy Exportable `
-HashAlgorithm sha256 -KeyLength 2048 `
-CertStoreLocation "Cert:\CurrentUser\My" -NotAfter (Get-Date).AddMonths(39)`
-Signer $cert -TextExtension @("2.5.29.37={text}1.3.6.1.5.5.7.3.2")

# Create Export Path
New-Item "$ENV:USERPROFILE\Documents\$path" -itemType Directory

# Export Root Cert to Path
$base64certificate = @"
-----BEGIN CERTIFICATE-----
$([Convert]::ToBase64String($cert.Export('Cert'), [System.Base64FormattingOptions]::InsertLineBreaks)))
-----END CERTIFICATE-----
"@

Set-Content -Path "$ENV:USERPROFILE\Documents\$path\P2SRootCert-$nickname.cer" -Value $base64certificate

# Export Child Cert to Path

Export-PfxCertificate -cert $cert2 -Password $securepassword -CryptoAlgorithmOption AES256_SHA256 -FilePath "$ENV:USERPROFILE\Documents\$path\P2SChildCert-$nickname.pfx" 

# Prompt for Azure
Write-Host "Check $ENV:USERPROFILE\Documents\$path\ for your new certificates.
If you wish to sign into Azure and upload the public Root cert automatically press enter. Otherwise, close out of the script" -ForegroundColor Cyan
Read-Host "Press Enter to upload the script to Azure, close out to exit..."

# Connect to Azure
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
Install-Module -Name Az -Scope CurrentUser -Repository PSGallery -Force
Connect-AzAccount

# Filter Resource Group
Read-Host "Press Enter to view a list of all the resource groups, take note of the resource group that contains your Virtual Network Gateway"
$getrg = Get-AzureResourceGroup | Format-Table -Property ResourceGroupName
$rg = Read-Host "Enter the name of the Resource Group where the Virtual Network Gateway lives"
Start-Sleep 1

# Filter Network Gateways
Read-Host "Press Enter to view a list of all the Network Gateways in the $rg group"
Get-AzVirtualNetworkGateway -Name * -ResourceGroupName "$rg" |Format-Table -Property Name,ResourceGroupName
$vn = Read-Host "Enter the name of the Virtual Network Gateway that needs the new certificate added"
Start-Sleep 1

# Update Root Cert
Write-Host "You are now going to update the Root Cert for the $vn Virtual Network Gateway." -ForegroundColor Cyan
Read-Host "Press Enter to update the certificate, close out to force end the script"
$CertificateText = for ($i=1; $i -lt $base64certificate.Length -1 ; $i++){$base64certificate[$i]}
Add-AzVpnClientRootCertificate -PublicCertData $CertificateText -ResourceGroupName "$rg" -VirtualNetworkGatewayName "$vn" -VpnClientRootCertificateName "P2sRootCert-$nickname"

Write-Host "Please verify that the VPN Root Cert is reflecting in Azure" -ForegroundColor Cyan
Get-AzVpnClientRootCertificate -VirtualNetworkGatewayName "$vn" -resourcegroupname "$rg" |format-table -Property Name,ProvisioningState,PublicCertData,ID
Start-Sleep 2

Read-Host "Press Enter to end the Azure connection"
Disconnect-AzAccount
Write-Host "Ending in 3 seconds"
Start-Sleep 3
Exit

