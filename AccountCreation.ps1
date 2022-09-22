<# 
#   Synopsis
This script was designed to help a company create user accounts in a hybrid environment using New-RemoteMailbox. 
New-Remotemailbox enables a mailbox object onto the AD account. For normal mail flow, this isn't necessary. But to receive
Scan-To-Email emails, it is necessary to have this enabled. Otherwise the SMTP traffic will not reach the user.

Currently it uses a static UPN, so it is not transferable to other companies.

#>
#Startup
$Host.UI.RawUI.WindowTitle = "Is This Thing On?"
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn


#User Details
$first = Read-Host 'First Name:'
$last = Read-Host 'Last Name:'
$un = Read-Host 'Username:'
$pw = Read-Host -AsSecureString 'Password (Must meet complexity requirements):'
$Name = $first + ' ' + $last
$choice = Read-Host -Prompt "Primary SMTP Address: Select 1 for @3ls.com, Select 2 for @omnivisions.com, Select 3 for omnicommunityhealth.com"

#Set the primary domain
$domain
    if ($choice -eq 1) {
          $domain = "3ls"
         }
       Elseif ($choice -eq 2) {
          $domain = "omnivisions"
         }
      Elseif ($choice -eq 3) {
          $domain = "omnicommunityhealth"
         }
      Else {
           write-host "This is not an option - exiting program."
               exit    
    }

#Account Creation
New-RemoteMailbox -Name "$first $last" -Password $pw -UserPrincipalName $un@$domain.com -PrimarySmtpAddress $un@$domain.com -RemoteRoutingAddress $un@$domain.mail.onmicrosoft.com

#Update First & Last Name fields in AD
Set-ADUser -Identity $un -GivenName $first
Set-ADUser -Identity $un -Surname $last
    
#Continue the account creation process 
Read-Host -Prompt "If no errors occured, then the account has been created. Please move the new account from AD/Users to the proper OU, and then press enter."

#Run Delta Sync
 $doSync = Read-Host -Prompt "Press 1 to run DeltaSync, press anything else to quit."
    if ($doSync -eq 1) {
        Write-Host "Syncing..."
        try {Start-ADSyncSyncCycle -PolicyType Delta | Out-Null}
        catch {
            $date = Get-Date
            $i = 1
            Do {
                $ADSyncService = Get-Service -Name ADSync
                $result = Get-EventLog application -Source "Directory Synchronization" -After $date `
                    | Where-Object {$_.EventID -eq "904"} `
                    | Where-Object {$_.Message -like "Import/Sync/Export cycle completed (Delta)."} -ErrorAction SilentlyContinue
                if (!$result){
                    if (("Running" -ne $ADSyncService.status) -or ($i -ge 45)) {
                        Write-Warning "Service ADSync is stuck. Elevated permissions are required to restart the service:"
                        Read-Host "Enter to continue..."
                        Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass `"Restart-Service -Name ADSync -Force`"" -Verb RunAs -Wait
                        Write-Output "ADSync Service should be restarted. Trying DeltaSync again..."
                        $i = 2
                    } else {
                        if ($i -eq 1) {Write-Host "Looks like there's already a sync running. Trying again..."}
                        if ($i % 5  -eq 0) {write-host "Still trying..."}
                        Sleep 1
                        $i++
                        Continue
                    }
                }
            } while (!$result)
            Start-ADSyncSyncCycle -PolicyType Delta | Out-Null
        }
        Write-Host "Success!"
        Start-Sleep 1
}
Else {
    Write-Host "Bye."
        exit    
}
Read-Host -Prompt "Please enter the Ticket # for this account creation request" | Out-File -Append c:\results.txt