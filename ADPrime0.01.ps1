Add-Type -AssemblyName PresentationFramework

function Create-MainWindow {
    $window = New-Object system.windows.window
    $window.Title = "Account Administration Toolkit"
    $window.Width = 800
    $window.Height = 600
    $window.Background = 'Black'

    $grid = New-Object system.windows.controls.grid
    $grid.Margin = '10'
    $window.Content = $grid

    $rowDef1 = New-Object system.windows.controls.rowdefinition
    $rowDef1.Height = "Auto"
    $rowDef2 = New-Object system.windows.controls.rowdefinition
    $rowDef2.Height = "*"
    $grid.RowDefinitions.Add($rowDef1)
    $grid.RowDefinitions.Add($rowDef2)

    $authGrid = New-Object system.windows.controls.grid
    $authGrid.Margin = '0,0,0,20'
    [void]$authGrid.ColumnDefinitions.Add((New-Object system.windows.controls.columndefinition))
    [void]$authGrid.ColumnDefinitions.Add((New-Object system.windows.controls.columndefinition))

    $connectButton = New-Object system.windows.controls.button
    $connectButton.Content = "Connect"
    $connectButton.Width = 100
    $connectButton.Height = 30
    $connectButton.Margin = '10'
    $connectButton.HorizontalAlignment = 'Left'
    $connectButton.VerticalAlignment = 'Top'
    $connectButton.Background = 'Gray'
    $connectButton.Foreground = 'White'
    $connectButton.Add_Click({ Connect-ToEntra })

    $statusLabel = New-Object system.windows.controls.label
    $statusLabel.Content = "Status: Disconnected"
    $statusLabel.Margin = '10'
    $statusLabel.Foreground = 'White'
    $statusLabel.HorizontalAlignment = 'Left'
    $statusLabel.VerticalAlignment = 'Top'
    
    $tenantNameLabel = New-Object system.windows.controls.label
    $tenantNameLabel.Content = "Tenant Name: "
    $tenantNameLabel.Margin = '10'
    $tenantNameLabel.Foreground = 'White'
    $tenantNameLabel.HorizontalAlignment = 'Left'
    $tenantNameLabel.VerticalAlignment = 'Top'

    $tenantIDLabel = New-Object system.windows.controls.label
    $tenantIDLabel.Content = "Tenant ID: "
    $tenantIDLabel.Margin = '10'
    $tenantIDLabel.Foreground = 'White'
    $tenantIDLabel.HorizontalAlignment = 'Left'
    $tenantIDLabel.VerticalAlignment = 'Top'

    $authGrid.Children.Add($connectButton)
    $authGrid.Children.Add($statusLabel)
    $authGrid.Children.Add($tenantNameLabel)
    $authGrid.Children.Add($tenantIDLabel)
    [void][System.Windows.Controls.Grid]::SetColumn($statusLabel, 1)
    [void][System.Windows.Controls.Grid]::SetColumn($tenantNameLabel, 1)
    [void][System.Windows.Controls.Grid]::SetColumn($tenantIDLabel, 1)
    [void][System.Windows.Controls.Grid]::SetRow($tenantNameLabel, 1)
    [void][System.Windows.Controls.Grid]::SetRow($tenantIDLabel, 2)

    $grid.Children.Add($authGrid)
    [void][System.Windows.Controls.Grid]::SetRow($authGrid, 0)

    # Add other tool sections below
    # Example: Create another grid for tools, add buttons and their respective functionalities

    $window.ShowDialog() | Out-Null
}

function Install-AzureADModule {
    Write-Output "Installing AzureAD module..."
    Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process -Force
    Install-Module -Name AzureAD -Force -Scope CurrentUser
}

function Connect-ToEntra {
    try {
        if (!(Get-Module -ListAvailable -Name AzureAD)) {
            Install-AzureADModule
        }
        
        Import-Module AzureAD
        Connect-AzureAD
        
        $statusLabel.Content = "Status: Connected"
        $tenant = Get-AzureADTenantDetail
        $tenantNameLabel.Content = "Tenant Name: " + $tenant.DisplayName
        $tenantIDLabel.Content = "Tenant ID: " + $tenant.ObjectId

        # Update button to disconnect
        $connectButton.Content = "Disconnect"
        $connectButton.Remove_Click({ Connect-ToEntra })
        $connectButton.Add_Click({ Disconnect-FromEntra })
    } catch {
        $statusLabel.Content = "Status: Disconnected - Error"
        Write-Output $_.Exception.Message
    }
}

function Disconnect-FromEntra {
    try {
        # Disconnect logic here, for AzureAD use Remove-AzureADServicePrincipal or another appropriate cmdlet
        $statusLabel.Content = "Status: Disconnected"
        $tenantNameLabel.Content = "Tenant Name: "
        $tenantIDLabel.Content = "Tenant ID: "

        # Update button to connect
        $connectButton.Content = "Connect"
        $connectButton.Remove_Click({ Disconnect-FromEntra })
        $connectButton.Add_Click({ Connect-ToEntra })
    } catch {
        $statusLabel.Content = "Status: Disconnected - Error"
        Write-Output $_.Exception.Message
    }
}

Create-MainWindow
