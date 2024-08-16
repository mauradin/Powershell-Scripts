# Load the necessary .NET assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create a form
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Azure AD Join Tool'
$form.Size = New-Object System.Drawing.Size(500, 400)
$form.StartPosition = 'CenterScreen'
$form.BackColor = [System.Drawing.Color]::FromArgb(51, 51, 51)

# Create and configure the status button
$btnStatus = New-Object System.Windows.Forms.Button
$btnStatus.Text = 'Join Status'
$btnStatus.Size = New-Object System.Drawing.Size(150, 30)
$btnStatus.Location = New-Object System.Drawing.Point(10, 10)
$btnStatus.BackColor = [System.Drawing.Color]::FromArgb(85, 85, 85)
$btnStatus.ForeColor = [System.Drawing.Color]::White

# Create and configure the join button
$btnJoin = New-Object System.Windows.Forms.Button
$btnJoin.Text = 'Join to Entra'
$btnJoin.Size = New-Object System.Drawing.Size(150, 30)
$btnJoin.Location = New-Object System.Drawing.Point(10, 50)
$btnJoin.BackColor = [System.Drawing.Color]::FromArgb(85, 85, 85)
$btnJoin.ForeColor = [System.Drawing.Color]::White

# Create and configure the status text box
$txtStatus = New-Object System.Windows.Forms.TextBox
$txtStatus.Multiline = $true
$txtStatus.ScrollBars = 'Vertical'
$txtStatus.Size = New-Object System.Drawing.Size(460, 300)
$txtStatus.Location = New-Object System.Drawing.Point(10, 90)
$txtStatus.BackColor = [System.Drawing.Color]::FromArgb(34, 34, 34)
$txtStatus.ForeColor = [System.Drawing.Color]::White
$txtStatus.ReadOnly = $true

# Add controls to the form
$form.Controls.Add($btnStatus)
$form.Controls.Add($btnJoin)
$form.Controls.Add($txtStatus)

# Variables for Azure AD join
$TenantId = "placeholder_tenant_id"
$ClientId = "placeholder_client_id"
$ClientSecret = ConvertTo-SecureString "placeholder_client_secret_id" -AsPlainText -Force
$ComputerName = $env:COMPUTERNAME

# Function to get and display the join status
function Get-JoinStatus {
    try {
        # Run dsregcmd /status and capture the output
        $output = & dsregcmd /status
        $txtStatus.Text = $output
    } catch {
        $txtStatus.Text = "Error running dsregcmd /status: $_"
    }
}

# Function to check if the computer is part of a domain
function Is-ComputerInDomain {
    try {
        $domainStatus = & dsregcmd /status
        if ($domainStatus -match "AzureADJoined\s*:\s*YES") {
            return $false  # Already joined to Azure AD
        } elseif ($domainStatus -match "DomainJoined\s*:\s*YES") {
            return $true   # Computer is part of a domain
        } else {
            return $false  # Not part of any domain
        }
    } catch {
        $txtStatus.Text = "Error checking domain status: $_"
        return $false
    }
}

# Function to unjoin from Active Directory Domain using WMIC
function Unjoin-FromDomain {
    try {
        # Run WMIC command to unjoin from domain and join a workgroup
        $txtStatus.Text = "Unjoining the computer from the domain and joining the workgroup..."
        Start-Process -NoNewWindow -FilePath "wmic.exe" -ArgumentList '/interactive:off ComputerSystem Where "Name=`"$ComputerName`"" Call UnJoinDomainOrWorkgroup FUnjoinOptions=0' -Wait
        Start-Process -NoNewWindow -FilePath "wmic.exe" -ArgumentList '/interactive:off ComputerSystem Where "Name=`"$ComputerName`"" Call JoinDomainOrWorkgroup name="WORKGROUP"' -Wait
        Start-Sleep -Seconds 60  # Wait for the computer to restart and fully unjoin
        $txtStatus.Text += "`nSuccessfully unjoined from the domain and joined workgroup."
    } catch {
        $txtStatus.Text += "`nFailed to unjoin from the domain or join workgroup: $_"
    }
}

# Function to join Microsoft Entra (Azure AD)
function Join-ToEntra {
    param (
        [string]$TenantId,
        [string]$ClientId,
        [securestring]$ClientSecret,
        [string]$ComputerName
    )

    try {
        # Install AzureAD module if not already installed
        if (-not (Get-Module -ListAvailable -Name AzureAD)) {
            Install-Module -Name AzureAD -Force
        }

        # Import the AzureAD module
        Import-Module AzureAD

        # Authenticate to Azure AD
        $secureAppCred = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.ClientCredential($ClientId, $ClientSecret)
        $authContext = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext("https://login.microsoftonline.com/$TenantId")
        $authResult = $authContext.AcquireTokenAsync("https://graph.microsoft.com/", $secureAppCred).Result

        $token = $authResult.AccessToken

        # Join the computer to Azure AD
        $txtStatus.Text = "Joining the computer to Microsoft Entra (Azure AD)..."
        $joinResult = Add-AzureADDevice -DeviceId $ComputerName -AccessToken $token

        if ($joinResult) {
            $txtStatus.Text += "`nSuccessfully joined to Microsoft Entra (Azure AD)."
        } else {
            $txtStatus.Text += "`nFailed to join Microsoft Entra (Azure AD)."
        }
    } catch {
        $txtStatus.Text = "Failed to join Microsoft Entra (Azure AD): $_"
    }
}

# Function to handle the Azure AD join process
function Handle-JoinToEntra {
    try {
        if (Is-ComputerInDomain) {
            # Unjoin from domain
            Unjoin-FromDomain
        }

        # Join to Azure AD
        Join-ToEntra -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret -ComputerName $ComputerName
    } catch {
        $txtStatus.Text = "An error occurred during the Azure AD join process: $_"
    }
}

# Attach the event handlers to the buttons
$btnStatus.Add_Click({ Get-JoinStatus })
$btnJoin.Add_Click({ Handle-JoinToEntra })

# Show the form
[void]$form.ShowDialog()
