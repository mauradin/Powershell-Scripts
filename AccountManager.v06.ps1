Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
Import-Module ActiveDirectory
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

function OUSelectorFunction {
    $dc_hash = @{}
    $selected_ou = $null

    $forest = Get-ADForest
    [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") | Out-Null
    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null

    function Get-NodeInfo($sender, $dn_textbox)
    {
        $selected_node = $sender.Node
        $dn_textbox.Text = $selected_node.Name
    }

    function Add-ChildNodes($sender)
    {
        $expanded_node = $sender.Node

        if ($expanded_node.Name -eq "root") {
            return
        }

        $expanded_node.Nodes.Clear() | Out-Null

        $dc_hostname = $dc_hash[$($expanded_node.Name -replace '(OU=[^,]+,)*((DC=\w+,?)+)','$2')]
        $child_OUs = Get-ADObject -Server $dc_hostname -Filter 'ObjectClass -eq "organizationalUnit" -or ObjectClass -eq "container"' -SearchScope OneLevel -SearchBase $expanded_node.Name
        if($child_OUs -eq $null) {
            $sender.Cancel = $true
        } else {
            foreach($ou in $child_OUs) {
                $ou_node = New-Object Windows.Forms.TreeNode
                $ou_node.Text = $ou.Name
                $ou_node.Name = $ou.DistinguishedName
                $ou_node.Nodes.Add('') | Out-Null
                $expanded_node.Nodes.Add($ou_node) | Out-Null
            }
        }
    }

    function Add-ForestNodes($forest, [ref]$dc_hash)
    {
        $ad_root_node = New-Object Windows.Forms.TreeNode
        $ad_root_node.Text = $forest.RootDomain
        $ad_root_node.Name = "root"
        $ad_root_node.Expand()

        $i = 1
        foreach ($ad_domain in $forest.Domains) {
            Write-Progress -Activity "Querying AD forest for domains and hostnames..." -Status $ad_domain -PercentComplete ($i++ / $forest.Domains.Count * 100)
            $dc = Get-ADDomainController -Server $ad_domain
            $dn = $dc.DefaultPartition
            $dc_hash.Value.Add($dn, $dc.Hostname)
            $dc_node = New-Object Windows.Forms.TreeNode
            $dc_node.Name = $dn
            $dc_node.Text = $dc.Domain
            $dc_node.Nodes.Add("") | Out-Null
            $ad_root_node.Nodes.Add($dc_node) | Out-Null
        }

        return $ad_root_node
    }
            
            $main_dlg_box                       = New-Object System.Windows.Forms.Form
            $main_dlg_box.ClientSize            = New-Object System.Drawing.Size(400,600)
            $main_dlg_box.MaximizeBox           = $false
            $main_dlg_box.MinimizeBox           = $false
            $main_dlg_box.FormBorderStyle       = 'FixedSingle'
            $main_dlg_box.BackColor             = [System.Drawing.ColorTranslator]::FromHtml("#252525")

            # widget size and location variables
            $ctrl_width_col = $main_dlg_box.ClientSize.Width/20
            $ctrl_height_row = $main_dlg_box.ClientSize.Height/15
            $max_ctrl_width = $main_dlg_box.ClientSize.Width - $ctrl_width_col*2
            $max_ctrl_height = $main_dlg_box.ClientSize.Height - $ctrl_height_row
            $right_edge_x = $max_ctrl_width
            $left_edge_x = $ctrl_width_col
            $bottom_edge_y = $max_ctrl_height
            $top_edge_y = $ctrl_height_row

            # setup text box showing the distinguished name of the currently selected node
            $dn_text_box = New-Object System.Windows.Forms.TextBox
            # can not set the height for a single line text box, that's controlled by the font being used
            $dn_text_box.Width = (14 * $ctrl_width_col)
            $dn_text_box.Location = New-Object System.Drawing.Point($left_edge_x, ($bottom_edge_y - $dn_text_box.Height))
            $dn_text_box.BackColor  = [System.Drawing.ColorTranslator]::FromHtml("#a3a3a3")
            $main_dlg_box.Controls.Add($dn_text_box)
            # /text box for dN

            # setup Ok button
            $ok_button = New-Object System.Windows.Forms.Button
            $ok_button.Size = New-Object System.Drawing.Size(($ctrl_width_col * 2), $dn_text_box.Height)
            $ok_button.Location = New-Object System.Drawing.Point(($right_edge_x - $ok_button.Width), ($bottom_edge_y - $ok_button.Height))
            $ok_button.Text = "Ok"
            $ok_button.DialogResult = 'OK'
            $ok_button.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
            $main_dlg_box.Controls.Add($ok_button)
            # /Ok button

            # setup tree selector showing the domains
            $ad_tree_view = New-Object System.Windows.Forms.TreeView
            $ad_tree_view.Size = New-Object System.Drawing.Size($max_ctrl_width, ($max_ctrl_height - $dn_text_box.Height - $ctrl_height_row*1.5))
            $ad_tree_view.Location = New-Object System.Drawing.Point($left_edge_x, $top_edge_y)
            $ad_tree_view.Nodes.Add($(Add-ForestNodes $forest ([ref]$dc_hash))) | Out-Null
            $ad_tree_view.Add_BeforeExpand({Add-ChildNodes $_})
            $ad_tree_view.Add_AfterSelect({Get-NodeInfo $_ $dn_text_box})
            $ok_button.Add_Click({$textDesiredOU.text = $dn_text_box.Text})
            $ad_tree_view.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#a3a3a3")
            $main_dlg_box.Controls.Add($ad_tree_view)
          # $textDesiredOU.text = $dn_text_box.Text
            # /tree selector

        $main_dlg_box.ShowDialog() | Out-Null

        return  $dn_text_box.Text
}



#Main Form
$formMain                        = New-Object system.Windows.Forms.Form
$formMain.ClientSize             = New-Object System.Drawing.Point(500,375)
$formMain.StartPosition          = 'CenterScreen'
$formMain.FormBorderStyle        = 'FixedSingle'
$formMain.MinimizeBox            = $false
$formMain.MaximizeBox            = $false
$formMain.ShowIcon               = $false
$formMain.text                   = "3ls Account Manager"
$formMain.TopMost                = $false
$formMain.BackColor              = [System.Drawing.ColorTranslator]::FromHtml("#252525")

#Panel Setup
$panelNewAccount                 = New-Object system.Windows.Forms.Panel
$panelNewAccount.height          = 120
$panelNewAccount.width           = 480
$panelNewAccount.Anchor          = 'top,right,left'
$panelNewAccount.location        = New-Object System.Drawing.Point(10,10)

$panelExistingAccount            = New-Object system.Windows.Forms.Panel
$panelExistingAccount.height     = 200
$panelExistingAccount.width      = 480
$panelExistingAccount.Anchor     = 'top,right,left'
$panelExistingAccount.location   = New-Object System.Drawing.Point(10,140)

#Tools List
$titleNew                        = New-Object system.Windows.Forms.Label
$titleNew.text                   = "NEW ACCOUNT OPTIONS"
$titleNew.AutoSize               = $true
$titleNew.width                  = 457
$titleNew.height                 = 142
$titleNew.Anchor                 = 'top,right,left'
$titleNew.location               = New-Object System.Drawing.Point(10,9)
$titleNew.Font                   = New-Object System.Drawing.Font('Consolas',15,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
$titleNew.ForeColor              = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

$formNewHACreation               = New-Object system.Windows.Forms.Button
$formNewHACreation.FlatStyle     = 'Flat'
$formNewHACreation.text          = "NEW HYBRID AD ACCOUNT"
$formNewHACreation.width         = 460
$formNewHACreation.height        = 30
$formNewHACreation.Anchor        = 'top,right,left'
$formNewHACreation.location      = New-Object System.Drawing.Point(10,40)
$formNewHACreation.Font          = New-Object System.Drawing.Font('Consolas',9)
$formNewHACreation.ForeColor     = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

$form365LicenseMgm               = New-Object system.Windows.Forms.Button
$form365LicenseMgm.FlatStyle     = 'Flat'
$form365LicenseMgm.text          = "365 LICENSE MANAGER"
$form365LicenseMgm.width         = 460
$form365LicenseMgm.height        = 30
$form365LicenseMgm.Anchor        = 'top,right,left'
$form365LicenseMgm.location      = New-Object System.Drawing.Point(10,80)
$form365LicenseMgm.Font          = New-Object System.Drawing.Font('Consolas',9)
$form365LicenseMgm.ForeColor     = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

$formOUMover                     = New-Object system.Windows.Forms.Button
$formOUMover.FlatStyle           = 'Flat'
$formOUMover.text                = "AD OU MOVER TOOL"
$formOUMover.width               = 225
$formOUMover.height              = 30
$formOUMover.Anchor              = 'top,right,left'
$formOUMover.location            = New-Object System.Drawing.Point(245,120)
$formOUMover.Font                = New-Object System.Drawing.Font('Consolas',9)
$formOUMover.ForeColor           = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

$titleExisting                   = New-Object system.Windows.Forms.Label
$titleExisting.text              = "EXISTING ACCOUNT OPTIONS"
$titleExisting.AutoSize          = $true
$titleExisting.width             = 457
$titleExisting.height            = 142
$titleExisting.Anchor            = 'top,right,left'
$titleExisting.location          = New-Object System.Drawing.Point(10,10)
$titleExisting.Font              = New-Object System.Drawing.Font('Consolas',15,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
$titleExisting.ForeColor         = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

$formUnlockAcc                   = New-Object system.Windows.Forms.Button
$formUnlockAcc.FlatStyle         = 'Flat'
$formUnlockAcc.text              = "UNLOCK / CHANGE PASSWORD"
$formUnlockAcc.width             = 225
$formUnlockAcc.height            = 30
$formUnlockAcc.Anchor            = 'top,right,left'
$formUnlockAcc.location          = New-Object System.Drawing.Point(10,40)
$formUnlockAcc.Font              = New-Object System.Drawing.Font('Consolas',9)
$formUnlockAcc.ForeColor         = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

$formRenameAcc                   = New-Object system.Windows.Forms.Button
$formRenameAcc.FlatStyle         = 'Flat'
$formRenameAcc.text              = "RENAME ACCOUNT"
$formRenameAcc.width             = 225
$formRenameAcc.height            = 30
$formRenameAcc.Anchor            = 'top,right,left'
$formRenameAcc.location          = New-Object System.Drawing.Point(245,40)
$formRenameAcc.Font              = New-Object System.Drawing.Font('Consolas',9)
$formRenameAcc.ForeColor         = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

$formGroupMembMang               = New-Object system.Windows.Forms.Button
$formGroupMembMang.FlatStyle     = 'Flat'
$formGroupMembMang.text          = "GROUP MEMBERSHIP MANAGER"
$formGroupMembMang.width         = 225
$formGroupMembMang.height        = 30
$formGroupMembMang.Anchor        = 'top,right,left'
$formGroupMembMang.location      = New-Object System.Drawing.Point(10,80)
$formGroupMembMang.Font          = New-Object System.Drawing.Font('Consolas',9)
$formGroupMembMang.ForeColor     = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

$formEnableRMBX                  = New-Object system.Windows.Forms.Button
$formEnableRMBX.FlatStyle        = 'Flat'
$formEnableRMBX.text             = "ENABLE REMOTE MAILBOX"
$formEnableRMBX.width            = 225
$formEnableRMBX.height           = 30
$formEnableRMBX.Anchor           = 'top,right,left'
$formEnableRMBX.location         = New-Object System.Drawing.Point(245,80)
$formEnableRMBX.Font             = New-Object System.Drawing.Font('Consolas',9)
$formEnableRMBX.ForeColor        = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

$formDisableAcc                  = New-Object system.Windows.Forms.Button
$formDisableAcc.FlatStyle        = 'Flat'
$formDisableAcc.text             = "DISABLE / ENABLE ACCOUNT"
$formDisableAcc.width            = 225
$formDisableAcc.height           = 30
$formDisableAcc.Anchor           = 'top,right,left'
$formDisableAcc.location         = New-Object System.Drawing.Point(10,120)
$formDisableAcc.Font             = New-Object System.Drawing.Font('Consolas',9)
$formDisableAcc.ForeColor        = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

$RunADSync                       = New-Object system.Windows.Forms.Button
$RunADSync.FlatStyle             = 'Flat'
$RunADSync.text                  = "RUN AD SYNC"
$RunADSync.width                 = 460
$RunADSync.height                = 30
$RunADSync.Anchor                = 'top,right,left'
$RunADSync.location              = New-Object System.Drawing.Point(10,170)
$RunADSync.Font                  = New-Object System.Drawing.Font('Consolas',9)
$RunADSync.ForeColor             = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

$formMain.controls.AddRange(@($panelNewAccount,$panelExistingAccount))
$panelNewAccount.controls.AddRange(@($titleNew,$formNewHACreation,$form365LicenseMgm))
$panelExistingAccount.controls.AddRange(@($titleExisting,$formUnlockAcc,$RunADSync,$formDisableAcc,$formRenameAcc,$formGroupMembMang,$formEnableRMBX,$formOUMover))

$ReceiptsFolder = ".\AccountManager\"
If (Test-Path $ReceiptsFolder) {
    Write-Host "${ReceiptsFolder} already exists."
}
Else {
    Write-Host "${ReceiptsFolder} does not exist. Creating a running text file for record keeping"
    Start-Sleep 1
    New-Item -Path "${ReceiptsFolder}" -ItemType Directory
    Write-Host "${ReceiptsFolder} was successfully created."
}

Start-Transcript -OutputDirectory "${ReceiptsFolder}"

#New Hybrid AD Account Form (COMPLETED)
    $formNewHACreation.Add_Click( {
        $formNewHA                      = New-Object System.Windows.Forms.Form
        $formNewHA.ClientSize           = New-Object System.Drawing.Point(580,570)
        $formNewHA.StartPosition        = 'CenterScreen'
        $formNewHA.FormBorderStyle      = 'FixedSingle'
        $formNewHA.MinimizeBox          = $false
        $formNewHA.MaximizeBox          = $false
        $formNewHA.ShowIcon             = $false
        $formNewHA.Text                 = "Create New Hybrid AD Account"
        $formNewHA.TopMost              = $false
        $formNewHA.AutoScroll           = $false
        $formNewHA.BackColor            = [System.Drawing.ColorTranslator]::FromHtml("#252525")

#Panel Setup
        $NewHybridADPanel               = New-Object System.Windows.Forms.Panel
        $NewHybridADPanel.height        = 510
        $NewHybridADPanel.width         = 572
        $NewHybridADPanel.Anchor        = 'top,right,left'
        $NewHybridADPanel.location      = New-Object System.Drawing.Point(10,10)
        $NewHybridADPanel.BackColor     = [System.Drawing.ColorTranslator]::FromHtml("#252525")

#Hybrid AD Account Panel
        $titleHybridFirstN              = New-Object system.Windows.Forms.Label
        $titleHybridFirstN.text         = "FIRST NAME"
        $titleHybridFirstN.AutoSize     = $true
        $titleHybridFirstN.width        = 457
        $titleHybridFirstN.height       = 142
        $titleHybridFirstN.Anchor       = 'top,right,left'
        $titleHybridFirstN.location     = New-Object System.Drawing.Point(5,10)
        $titleHybridFirstN.Font         = New-Object System.Drawing.Font('Consolas',15,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
        $titleHybridFirstN.ForeColor    = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

        $titleHybridLastN               = New-Object system.Windows.Forms.Label
        $titleHybridLastN.text          = "LAST NAME"
        $titleHybridLastN.AutoSize      = $true
        $titleHybridLastN.width         = 457
        $titleHybridLastN.height        = 142
        $titleHybridLastN.Anchor        = 'top,right,left'
        $titleHybridLastN.location      = New-Object System.Drawing.Point(255,10)
        $titleHybridLastN.Font          = New-Object System.Drawing.Font('Consolas',15,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
        $titleHybridLastN.ForeColor     = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

        $textNewHAFirstN                = New-Object System.Windows.Forms.TextBox
        $textNewHAFirstN.width          = 230
        $textNewHAFirstN.height         = 30
        $textNewHAFirstN.Anchor         = 'top,right,left'
        $textNewHAFirstN.location       = New-Object System.Drawing.Point(10,40)
        $textNewHAFirstN.Font           = New-Object System.Drawing.Font('Consolas',9)
        $textNewHAFirstN.ForeColor      = [System.Drawing.ColorTranslator]::FromHtml("#000000")

        $textNewHALastN                 = New-Object System.Windows.Forms.TextBox
        $textNewHALastN.width           = 230
        $textNewHALastN.height          = 30
        $textNewHALastN.Anchor          = 'top,right,left'
        $textNewHALastN.location        = New-Object System.Drawing.Point(260,40)
        $textNewHALastN.Font            = New-Object System.Drawing.Font('Consolas',9)
        $textNewHALastN.ForeColor       = [System.Drawing.ColorTranslator]::FromHtml("#000000")

        $titleHybridUserName            = New-Object system.Windows.Forms.Label
        $titleHybridUserName.text       = "USER NAME"
        $titleHybridUserName.AutoSize   = $true
        $titleHybridUserName.width      = 457
        $titleHybridUserName.height     = 142
        $titleHybridUserName.Anchor     = 'top,right,left'
        $titleHybridUserName.location   = New-Object System.Drawing.Point(5,70)
        $titleHybridUserName.Font       = New-Object System.Drawing.Font('Consolas',15,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
        $titleHybridUserName.ForeColor  = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

        $textNewHAUsername              = New-Object System.Windows.Forms.TextBox
        $textNewHAUsername.width        = 230
        $textNewHAUsername.height       = 30
        $textNewHAUsername.Anchor       = 'top,right,left'
        $textNewHAUsername.location     = New-Object System.Drawing.Point(10,100)
        $textNewHAUsername.Font         = New-Object System.Drawing.Font('Consolas',9)
        $textNewHAUsername.ForeColor    = [System.Drawing.ColorTranslator]::FromHtml("#000000")

        $titleDomainName                = New-Object system.Windows.Forms.Label
        $titleDomainName.text           = "EMAIL DOMAIN"
        $titleDomainName.AutoSize       = $true
        $titleDomainName.width          = 457
        $titleDomainName.height         = 142
        $titleDomainName.Anchor         = 'top,right,left'
        $titleDomainName.location       = New-Object System.Drawing.Point(255,70)
        $titleDomainName.Font           = New-Object System.Drawing.Font('Consolas',15,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
        $titleDomainName.ForeColor      = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

        $listNewHADomainList            = New-Object System.Windows.Forms.ComboBox
        $listNewHADomainList.width      = 230
        $listNewHADomainList.height     = 30
        $listNewHADomainList.Anchor     = 'top,right,left'
        $listNewHADomainList.location   = New-Object System.Drawing.Point(260,100)
        $listNewHADomainList.Font       = New-Object System.Drawing.Font('Consolas',9)
        $listNewHADomainList.ForeColor  = [System.Drawing.ColorTranslator]::FromHtml("#000000")
      #  @('@omnivisions.com','@3ls.com','@omnicommunityhealth.com','@theomnifamily.com','@omnicareinstitute.com') | ForEach-Object {[void] $listNewHADomainList.Items.Add($_)}
        (Get-ADForest).UPNSuffixes | ForEach-Object {[void] $listNewHADomainList.Items.Add($_)}

        $titleHybridPassword            = New-Object system.Windows.Forms.Label
        $titleHybridPassword.text       = "PASSWORD"
        $titleHybridPassword.AutoSize   = $true
        $titleHybridPassword.width      = 457
        $titleHybridPassword.height     = 142
        $titleHybridPassword.Anchor     = 'top,right,left'
        $titleHybridPassword.location   = New-Object System.Drawing.Point(5,130)
        $titleHybridPassword.Font       = New-Object System.Drawing.Font('Consolas',15,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
        $titleHybridPassword.ForeColor  = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

        $textNewHAPassword              = New-Object System.Windows.Forms.MaskedTextBox
        $textNewHAPassword.width        = 230
        $textNewHAPassword.height       = 30
        $textNewHAPassword.Anchor       = 'top,right,left'
        $textNewHAPassword.location     = New-Object System.Drawing.Point(10,160)
        $textNewHAPassword.Font         = New-Object System.Drawing.Font('Consolas',9)
        $textNewHAPassword.ForeColor    = [System.Drawing.ColorTranslator]::FromHtml("#000000")
        $textNewHAPassword.PasswordChar = "*"

        $titleHybridAcctTicket          = New-Object system.Windows.Forms.Label
        $titleHybridAcctTicket.text     = "TICKET #"
        $titleHybridAcctTicket.AutoSize = $true
        $titleHybridAcctTicket.width    = 457
        $titleHybridAcctTicket.height   = 142
        $titleHybridAcctTicket.Anchor   = 'top,right,left'
        $titleHybridAcctTicket.location = New-Object System.Drawing.Point(255,130)
        $titleHybridAcctTicket.Font     = New-Object System.Drawing.Font('Consolas',15,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
        $titleHybridAcctTicket.ForeColor= [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

        $textNewHATicketNumb            = New-Object System.Windows.Forms.TextBox
        $textNewHATicketNumb.width      = 230
        $textNewHATicketNumb.height     = 30
        $textNewHATicketNumb.Anchor     = 'top,right,left'
        $textNewHATicketNumb.location   = New-Object System.Drawing.Point(260,160)
        $textNewHATicketNumb.Font       = New-Object System.Drawing.Font('Consolas',9)
        $textNewHATicketNumb.ForeColor  = [System.Drawing.ColorTranslator]::FromHtml("#000000")

        $HybridAcctLogs                 = New-Object System.Windows.Forms.RichTextBox
        $HybridAcctLogs.width           = 560
        $HybridAcctLogs.height          = 200
        $HybridAcctLogs.location        = New-Object System.Drawing.Point(10,350)
        $HybridAcctLogs.Font            = New-Object System.Drawing.Font('Consolas',9)
        $HybridAcctLogs.BackColor       = [System.Drawing.ColorTranslator]::FromHtml("#252525")
        $HybridAcctLogs.ForeColor       = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
        $HybridAcctLogs.ReadOnly        = $true

        $buttonNewHACreate              = New-Object System.Windows.Forms.Button
        $buttonNewHACreate.FlatStyle    = 'Flat'
        $buttonNewHACreate.Text         = "CREATE HYBRID AD ACCOUNT"
        $buttonNewHACreate.width        = 560
        $buttonNewHACreate.height       = 30
        $buttonNewHACreate.Location     = New-Object System.Drawing.Point(10, 300)
        $buttonNewHACreate.Font         = New-Object System.Drawing.Font('Consolas',9)
        $buttonNewHACreate.ForeColor    = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

        $titleAccountOptions            = New-Object system.Windows.Forms.Label
        $titleAccountOptions.text       = "ACCOUNT OPTIONS"
        $titleAccountOptions.AutoSize   = $true
        $titleAccountOptions.width      = 457
        $titleAccountOptions.height     = 142
        $titleAccountOptions.Anchor     = 'top,right,left'
        $titleAccountOptions.location   = New-Object System.Drawing.Point(10,200)
        $titleAccountOptions.Font       = New-Object System.Drawing.Font('Consolas',15,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
        $titleAccountOptions.ForeColor  = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

        $optionLocalADOnly              = New-Object system.Windows.Forms.Checkbox
        $optionLocalADOnly.text         = "Local AD Account (No RemoteMailbox)"
        $optionLocalADOnly.AutoSize     = $true
        $optionLocalADOnly.width        = 457
        $optionLocalADOnly.height       = 142
        $optionLocalADOnly.Anchor       = 'top,right,left'
        $optionLocalADOnly.location     = New-Object System.Drawing.Point(260,230)
        $optionLocalADOnly.Font         = New-Object System.Drawing.Font('Consolas',8,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
        $optionLocalADOnly.ForeColor    = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
        #$optionLocalADOnly.Checked = $false

        $textDesiredOU                  = New-Object System.Windows.Forms.TextBox
        $textDesiredOU.width            = 480
        $textDesiredOU.height           = 30
        $textDesiredOU.Anchor           = 'top,right,left'
        $textDesiredOU.location         = New-Object System.Drawing.Point(10,270)
        $textDesiredOU.Font             = New-Object System.Drawing.Font('Consolas',9)
        $textDesiredOU.ForeColor        = [System.Drawing.ColorTranslator]::FromHtml("#000000")
        $textDesiredOU.ReadOnly         = $true

        $buttonDesiredNewOU             = New-Object System.Windows.Forms.Button
        $buttonDesiredNewOU.FlatStyle   = 'Flat'
        $buttonDesiredNewOU.Text        = "SELECT DESIRED OU"
        $buttonDesiredNewOU.width       = 225
        $buttonDesiredNewOU.height      = 30
        $buttonDesiredNewOU.Location    = New-Object System.Drawing.Point(10, 230)
        $buttonDesiredNewOU.Font        = New-Object System.Drawing.Font('Consolas',9)
        $buttonDesiredNewOU.ForeColor   = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

    #Hybrid Account Range Control
        $formNewHA.controls.AddRange(@($textNewHAFirstN,$textNewHALastN,$textNewHAUsername,$listNewHADomainList,$titleHybridAcctTicket,$textNewHAPassword,$textNewHATicketNumb,$titleHybridPassword,$titleDomainName,$titleHybridUserName,
        $titleHybridFirstN,$titleHybridLastN,$buttonDesiredNewOU,$optionLocalADOnly,$buttonNewHACreate,$HybridAcctLogs,$titleAccountOptions,$textDesiredOU,$NewHybridADPanel))

    #Set up Button Clicks
        $buttonDesiredNewOU.Add_Click({OUSelectorFunction})

        $buttonNewHACreate.Add_Click( 
        {
            #Text Field Variables
            $varNewHAFirstN                     =  $($textNewHAFirstN.text)
            $varNewHALastN                      =  $($textNewHALastN.text)
            $varNewHAUN                         =  $($textNewHAUsername.text)
            $varNewHAPW                         =  $($textNewHAPassword.text) | ConvertTo-SecureString -AsPlainText -Force
            $varNewHATicket                     =  $($textNewHATicketNumb.text)
            $varNewHADomain                     =  $($listNewHADomainList.text)
            $varDesiredOU                       =  $($textDesiredOU.text)
            $varNewHATime                       =  Get-Date

           #Checking if Hybrid Account already exists. If it does, prompt an error. If it does not, then create the account.                 
           $varNewHACheck = Get-ADUser -filter {sAMAccountName -eq $varNewHAUN}

        if ($optionLocalADOnly.Checked -eq $false){
            
            if ($varNewHACheck -ne $null) {
            #Display error results on screen
            $varNewHALog = "The account $varNewHAUN already exists. Please check in AD if the account is active or disabled."
            $HybridAcctLogs.text = $varNewHALog
           }
            else {
            New-RemoteMailbox -Name "$varNewHAFirstN $varNewHALastN" -Password $varNewHAPW -UserPrincipalName $varNewHAUN@$varNewHADomain -PrimarySmtpAddress $varNewHAUN@$varNewHADomain #-RemoteRoutingAddress $un@$domain.mail.onmicrosoft.com
            Start-Sleep -Seconds 1
            Set-ADUser -Identity $varNewHAUN -GivenName $varNewHAFirstN
            Start-Sleep -Seconds 1
            Set-ADUser -Identity $varNewHAUN -Surname $varNewHALastN
            Start-Sleep -Seconds 1
            Get-ADUser -Identity "$varNewHAUN" | Move-ADObject -TargetPath "$varDesiredOU"
            Start-Sleep -Seconds 1

            #Display success results on screen
            $varNewHALog = @" 
            "A New Hybrid Exchange account: $varNewHAUN@$varNewHADomain was created
            First Name: $varNewHAFirstN
            Last Name: $varNewHALastN
            Username: $varNewHAUN
            Ticket#: $varNewHATicket
            Date Created: $varNewHATime
"@
            $HybridAcctLogs.text = $varNewHALog
           }}

        else {                   

            if ($varNewHACheck -ne $null) {
            #Display error results on screen
            $varNewHALog = "The account $varNewHAUN already exists. Please check in AD if the account is active or disabled."
            $HybridAcctLogs.text = $varNewHALog
           }
            else {
            New-ADUser -Name "$varNewHAFirstN $varNewHALastN" -AccountPassword $varNewHAPW -UserPrincipalName $varNewHAUN@$varNewHADomain -SamAccountName $varNewHAUN -Enabled $true #-RemoteRoutingAddress $un@$domain.mail.onmicrosoft.com
            Start-Sleep -Seconds 1
            Set-ADUser -Identity $varNewHAUN -GivenName $varNewHAFirstN
            Start-Sleep -Seconds 1
            Set-ADUser -Identity $varNewHAUN -Surname $varNewHALastN
            Start-Sleep -Seconds 1
            Set-ADUser -Identity $varNewHAUN -DisplayName "$varNewHAFirstN $varNewHALastN"
            Start-Sleep -Seconds 1
            Get-ADUser -Identity "$varNewHAUN" | Move-ADObject -TargetPath "$varDesiredOU"
            Start-Sleep -Seconds 1


            #Display success results on screen
            $varNewHALog = @" 
            A New Local AD account: $varNewHAUN@$varNewHADomain was created
            First Name: $varNewHAFirstN
            Last Name: $varNewHALastN
            Username: $varNewHAUN
            Ticket#: $varNewHATicket
            Date Created: $varNewHATime
"@
            $HybridAcctLogs.text = $varNewHALog
           }}

           #Log & report externally
            If (Test-Path $ReceiptsFolder\NewHybridAccount) {
                Write-Host "Creating receipts folder..."
            }
            Else {
                Start-Sleep 1
                New-Item -Path "${ReceiptsFolder}\NewHybridAccount" -ItemType Directory
            }
            Out-File -FilePath $ReceiptsFolder\NewHybridAccount\$($varNewHAFirstN)_$($varNewHALastN)_$($varNewHAUN)_$($varNewHATicket).txt -Encoding utf8
            Add-Content -Path $ReceiptsFolder\NewHybridAccount\$($varNewHAFirstN)_$($varNewHALastN)_$($varNewHAUN)_$($varNewHATicket).txt "$varNewHALog"
        }
        )
        #End Button Click
        [void]$formNewHA.ShowDialog()
    })

    
#New 365 Account Form 
    $form365LicenseMgm.Add_Click( {
    })


#AD OU Mover Tool (COMPLETED)
    $formOUMover.Add_Click( { 
        $formADOUMover                  = New-Object System.Windows.Forms.Form
        $formADOUMover.ClientSize       = New-Object System.Drawing.Point(580,570)
        $formADOUMover.StartPosition    = 'CenterScreen'
        $formADOUMover.FormBorderStyle  = 'FixedSingle'
        $formADOUMover.MinimizeBox      = $false
        $formADOUMover.MaximizeBox      = $false
        $formADOUMover.ShowIcon         = $false
        $formADOUMover.Text             = "AD OU Mover Tool"
        $formADOUMover.TopMost          = $false
        $formADOUMover.AutoScroll       = $false
        $formADOUMover.BackColor        = [System.Drawing.ColorTranslator]::FromHtml("#252525")

    #Panel Setup
        $panelADOUMover                 = New-Object System.Windows.Forms.Panel
        $panelADOUMover.height          = 510
        $panelADOUMover.width           = 572
        $panelADOUMover.Anchor          = 'top,right,left'
        $panelADOUMover.location        = New-Object System.Drawing.Point(10,10)
        $panelADOUMover.BackColor       = [System.Drawing.ColorTranslator]::FromHtml("#252525")

    #Hybrid AD Account Panel
        $titleOUFirstName               = New-Object system.Windows.Forms.Label
        $titleOUFirstName.text          = "FIRST NAME"
        $titleOUFirstName.AutoSize      = $true
        $titleOUFirstName.width         = 457
        $titleOUFirstName.height        = 142
        $titleOUFirstName.Anchor        = 'top,right,left'
        $titleOUFirstName.location      = New-Object System.Drawing.Point(5,70)
        $titleOUFirstName.Font          = New-Object System.Drawing.Font('Consolas',15,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
        $titleOUFirstName.ForeColor     = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

        $titleOULastName                = New-Object system.Windows.Forms.Label
        $titleOULastName.text           = "LAST NAME"
        $titleOULastName.AutoSize       = $true
        $titleOULastName.width          = 457
        $titleOULastName.height         = 142
        $titleOULastName.Anchor         = 'top,right,left'
        $titleOULastName.location       = New-Object System.Drawing.Point(255,70)
        $titleOULastName.Font           = New-Object System.Drawing.Font('Consolas',15,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
        $titleOULastName.ForeColor      = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

        $textOUFirstName                = New-Object System.Windows.Forms.TextBox
        $textOUFirstName.width          = 230
        $textOUFirstName.height         = 30
        $textOUFirstName.Anchor         = 'top,right,left'
        $textOUFirstName.location       = New-Object System.Drawing.Point(10,100)
        $textOUFirstName.Font           = New-Object System.Drawing.Font('Consolas',9)
        $textOUFirstName.ForeColor      = [System.Drawing.ColorTranslator]::FromHtml("#000000")
        $textOUFirstName.ReadOnly       = $true
        $textOUFirstName.text           = $null

        $textOULastName                 = New-Object System.Windows.Forms.TextBox
        $textOULastName.width           = 230
        $textOULastName.height          = 30
        $textOULastName.Anchor          = 'top,right,left'
        $textOULastName.location        = New-Object System.Drawing.Point(260,100)
        $textOULastName.Font            = New-Object System.Drawing.Font('Consolas',9)
        $textOULastName.ForeColor       = [System.Drawing.ColorTranslator]::FromHtml("#000000")
        $textOULastName.ReadOnly        = $true

        $titleOUUsername                = New-Object system.Windows.Forms.Label
        $titleOUUsername.text           = "USER NAME"
        $titleOUUsername.AutoSize       = $true
        $titleOUUsername.width          = 457
        $titleOUUsername.height         = 142
        $titleOUUsername.Anchor         = 'top,right,left'
        $titleOUUsername.location       = New-Object System.Drawing.Point(10,10)
        $titleOUUsername.Font           = New-Object System.Drawing.Font('Consolas',15,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
        $titleOUUsername.ForeColor      = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

        $textOUUsername                 = New-Object System.Windows.Forms.TextBox
        $textOUUsername.width           = 230
        $textOUUsername.height          = 30
        $textOUUsername.Anchor          = 'top,right,left'
        $textOUUsername.location        = New-Object System.Drawing.Point(10,40)
        $textOUUsername.Font            = New-Object System.Drawing.Font('Consolas',9)
        $textOUUsername.ForeColor       = [System.Drawing.ColorTranslator]::FromHtml("#000000")

        $titleCurrentOU                 = New-Object system.Windows.Forms.Label
        $titleCurrentOU.text            = "CURRENT OU"
        $titleCurrentOU.AutoSize        = $true
        $titleCurrentOU.width           = 457
        $titleCurrentOU.height          = 142
        $titleCurrentOU.Anchor          = 'top,right,left'
        $titleCurrentOU.location        = New-Object System.Drawing.Point(5,130)
        $titleCurrentOU.Font            = New-Object System.Drawing.Font('Consolas',15,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
        $titleCurrentOU.ForeColor       = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

        $textCurrentOU                  = New-Object System.Windows.Forms.TextBox
        $textCurrentOU.width            = 480
        $textCurrentOU.height           = 30
        $textCurrentOU.Anchor           = 'top,right,left'
        $textCurrentOU.location         = New-Object System.Drawing.Point(10,160)
        $textCurrentOU.Font             = New-Object System.Drawing.Font('Consolas',9)
        $textCurrentOU.ForeColor        = [System.Drawing.ColorTranslator]::FromHtml("#000000")
        $textCurrentOU.ReadOnly         = $true

        $titleSelectNewOU              = New-Object system.Windows.Forms.Label
        $titleSelectNewOU.text         = "SELECT NEW OU"
        $titleSelectNewOU.AutoSize     = $true
        $titleSelectNewOU.width        = 457
        $titleSelectNewOU.height       = 142
        $titleSelectNewOU.Anchor       = 'top,right,left'
        $titleSelectNewOU.location     = New-Object System.Drawing.Point(5,190)
        $titleSelectNewOU.Font         = New-Object System.Drawing.Font('Consolas',15,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
        $titleSelectNewOU.ForeColor    = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

        $buttonSelectNewOU             = New-Object System.Windows.Forms.Button
        $buttonSelectNewOU.FlatStyle   = 'Flat'
        $buttonSelectNewOU.Text        = "SELECT NEW OU"
        $buttonSelectNewOU.width       = 225
        $buttonSelectNewOU.height      = 30
        $buttonSelectNewOU.Location    = New-Object System.Drawing.Point(10, 225)
        $buttonSelectNewOU.Font        = New-Object System.Drawing.Font('Consolas',9)
        $buttonSelectNewOU.ForeColor   = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

        $OUMoveToolLogs                 = New-Object System.Windows.Forms.RichTextBox
        $OUMoveToolLogs.width           = 560
        $OUMoveToolLogs.height          = 150
        $OUMoveToolLogs.location        = New-Object System.Drawing.Point(10,400)
        $OUMoveToolLogs.Font            = New-Object System.Drawing.Font('Consolas',9)
        $OUMoveToolLogs.BackColor       = [System.Drawing.ColorTranslator]::FromHtml("#252525")
        $OUMoveToolLogs.ForeColor       = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
        $OUMoveToolLogs.ReadOnly        = $true

        $buttonMoveOUTool               = New-Object System.Windows.Forms.Button
        $buttonMoveOUTool.FlatStyle     = 'Flat'
        $buttonMoveOUTool.Text          = "MOVE ACCOUNT TO NEW OU"
        $buttonMoveOUTool.width         = 560
        $buttonMoveOUTool.height        = 30
        $buttonMoveOUTool.Location      = New-Object System.Drawing.Point(10, 350)
        $buttonMoveOUTool.Font          = New-Object System.Drawing.Font('Consolas',9)
        $buttonMoveOUTool.ForeColor     = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

        $buttonOUChecker                = New-Object System.Windows.Forms.Button
        $buttonOUChecker.FlatStyle      = 'Flat'
        $buttonOUChecker.Text           = "CHECK OU"
        $buttonOUChecker.width          = 225
        $buttonOUChecker.height         = 30
        $buttonOUChecker.Location       = New-Object System.Drawing.Point(260, 35)
        $buttonOUChecker.Font           = New-Object System.Drawing.Font('Consolas',9)
        $buttonOUChecker.ForeColor      = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

        $textDesiredOU                  = New-Object System.Windows.Forms.TextBox
        $textDesiredOU.width            = 480
        $textDesiredOU.height           = 30
        $textDesiredOU.Anchor           = 'top,right,left'
        $textDesiredOU.location         = New-Object System.Drawing.Point(10,270)
        $textDesiredOU.Font             = New-Object System.Drawing.Font('Consolas',9)
        $textDesiredOU.ForeColor        = [System.Drawing.ColorTranslator]::FromHtml("#000000")
        $textDesiredOU.ReadOnly         = $true


    #Hybrid Account Range Control
        $formADOUMover.controls.AddRange(@($textDesiredOU,$textCurrentOU,$buttonOUChecker,$titleCurrentOU,$titleSelectNewOU,$buttonSelectNewOU,$titleOUUsername,
        $textOUUsername,$textOUFirstName,$textOULastName,$titleOUFirstName,$titleOULastName,$buttonMoveOUTool,$OUMoveToolLogs,$panelADOUMover))

    #Set up Button Clicks
        $buttonOUChecker.Add_Click( {
            #Text Field Variables
        
            $OUNamePull                      = Get-ADUser -Identity $textOUUsername.text
            $OUFirstNameResult               = $OUNamePull.GivenName
            $OULastNameResult                = $OUNamePull.surname
            $OUCurrentOU                     = $OUNamePull.DistinguishedName

            #Populate First/Last/Current fields
            $textOUFirstName.text            = $OUFirstNameResult
            $textOULastName.text             = $OULastNameResult
            $textCurrentOU.text              = $OUCurrentOU
        })
        #End OU Checker

        $buttonSelectNewOU.Add_Click({OUSelectorFunction})
        #end OU Selector

        $buttonMoveOUTool.Add_Click( 

            {
                #Text Field Variables
                $varOUFName                      =  $($textOUFirstName.text)
                $varOULName                      =  $($textOULastName.text)
                $varOUUser                       =  $($textOUUsername.text)
                $varCurrentOU                    =  $($textCurrentOU.text)
                $VarDesiredOU                    =  $($textDesiredOU.text)
                $varOUTime                       =  Get-Date

           #Checking if OU Check was ran, if not; error.                
           If (($textOUFirstName.text -lt 0) -And ($textDesiredOU.text.length -lt 0)) {
            #Display error results on screen
            $varOULogs = "Please enter a username, click on Check OU first, and make sure you have selected your desired OU!"
            $OUMoveToolLogs.text = $varOULogs
           }
           Else {
            #Move AD Object to new OU command
            Get-ADUser -Identity "$varOUUser" | Move-ADObject -TargetPath "$varDesiredOU"

            #Display success results on screen
            $varOULogs = @" 
            "An account was moved to a different OU
            First Name: $varOUFName
            Last Name: $varOULName
            Username: $varOUUser
            Old OU: $varCurrentOU
            New OU: $VarDesiredOU
            Date Created: $varOUTime
"@
            $OUMoveToolLogs.text = $varOULogs
           }

           #Log & report externally
            If (Test-Path $ReceiptsFolder\OUMoveTool) {
                Write-Host "Creating receipts folder..."
           }
            Else {
                Start-Sleep 1
                New-Item -Path "${ReceiptsFolder}\OUMoveTool" -ItemType Directory
           }
            Out-File -FilePath $ReceiptsFolder\OUMoveTool\$($varOUFName)_$($varOULName)_$($varOUUser).txt -Encoding utf8
            Add-Content -Path $ReceiptsFolder\OUMoveTool\$($varOUFName)_$($varOULName)_$($varOUUser).txt "$varOULogs"
        }
        )
        #End Move Tool
        [void]$formADOUMover.ShowDialog()
})

#Disable Account Tool (COMPLETED)
$formDisableAcc.Add_Click( { 
    $formEDAccount                   = New-Object System.Windows.Forms.Form
    $formEDAccount.ClientSize        = New-Object System.Drawing.Point(580,350)
    $formEDAccount.StartPosition     = 'CenterScreen'
    $formEDAccount.FormBorderStyle   = 'FixedSingle'
    $formEDAccount.MinimizeBox       = $false
    $formEDAccount.MaximizeBox       = $false
    $formEDAccount.ShowIcon          = $false
    $formEDAccount.Text              = "Enable / Disable Account"
    $formEDAccount.TopMost           = $false
    $formEDAccount.AutoScroll        = $false
    $formEDAccount.BackColor         = [System.Drawing.ColorTranslator]::FromHtml("#252525")

    #Panel Setup
    $panelEDAccount                  = New-Object System.Windows.Forms.Panel
    $panelEDAccount.height           = 350
    $panelEDAccount.width            = 572
    $panelEDAccount.Anchor           = 'top,right,left'
    $panelEDAccount.location         = New-Object System.Drawing.Point(10,10)
    $panelEDAccount.BackColor        = [System.Drawing.ColorTranslator]::FromHtml("#252525")

    #Enable Disable Panels
    $titleEDAccountFName             = New-Object system.Windows.Forms.Label
    $titleEDAccountFName.text        = "FIRST NAME"
    $titleEDAccountFName.AutoSize    = $true
    $titleEDAccountFName.width       = 457
    $titleEDAccountFName.height      = 142
    $titleEDAccountFName.Anchor      = 'top,right,left'
    $titleEDAccountFName.location    = New-Object System.Drawing.Point(5,70)
    $titleEDAccountFName.Font        = New-Object System.Drawing.Font('Consolas',15,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
    $titleEDAccountFName.ForeColor   = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

    $titleEDAccountLName             = New-Object system.Windows.Forms.Label
    $titleEDAccountLName.text        = "LAST NAME"
    $titleEDAccountLName.AutoSize    = $true
    $titleEDAccountLName.width       = 457
    $titleEDAccountLName.height      = 142
    $titleEDAccountLName.Anchor      = 'top,right,left'
    $titleEDAccountLName.location    = New-Object System.Drawing.Point(255,70)
    $titleEDAccountLName.Font        = New-Object System.Drawing.Font('Consolas',15,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
    $titleEDAccountLName.ForeColor   = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

    $textEDAccountFirstN             = New-Object System.Windows.Forms.TextBox
    $textEDAccountFirstN.width       = 230
    $textEDAccountFirstN.height      = 30
    $textEDAccountFirstN.Anchor      = 'top,right,left'
    $textEDAccountFirstN.location    = New-Object System.Drawing.Point(10,100)
    $textEDAccountFirstN.Font        = New-Object System.Drawing.Font('Consolas',9)
    $textEDAccountFirstN.ForeColor   = [System.Drawing.ColorTranslator]::FromHtml("#000000")
    $textEDAccountFirstN.ReadOnly    = $true
    $textEDAccountFirstN.text        = $null

    $textEDAccountLastN              = New-Object System.Windows.Forms.TextBox
    $textEDAccountLastN.width        = 230
    $textEDAccountLastN.height       = 30
    $textEDAccountLastN.Anchor       = 'top,right,left'
    $textEDAccountLastN.location     = New-Object System.Drawing.Point(260,100)
    $textEDAccountLastN.Font         = New-Object System.Drawing.Font('Consolas',9)
    $textEDAccountLastN.ForeColor    = [System.Drawing.ColorTranslator]::FromHtml("#000000")
    $textEDAccountLastN.ReadOnly     = $true

    $titleEDUsername                 = New-Object system.Windows.Forms.Label
    $titleEDUsername.text            = "USER NAME"
    $titleEDUsername.AutoSize        = $true
    $titleEDUsername.width           = 457
    $titleEDUsername.height          = 142
    $titleEDUsername.Anchor          = 'top,right,left'
    $titleEDUsername.location        = New-Object System.Drawing.Point(10,10)
    $titleEDUsername.Font            = New-Object System.Drawing.Font('Consolas',15,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
    $titleEDUsername.ForeColor       = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

    $textEDUsername                  = New-Object System.Windows.Forms.TextBox
    $textEDUsername.width            = 230
    $textEDUsername.height           = 30
    $textEDUsername.Anchor           = 'top,right,left'
    $textEDUsername.location         = New-Object System.Drawing.Point(10,40)
    $textEDUsername.Font             = New-Object System.Drawing.Font('Consolas',9)
    $textEDUsername.ForeColor        = [System.Drawing.ColorTranslator]::FromHtml("#000000")

    $titleEDEmail                    = New-Object system.Windows.Forms.Label
    $titleEDEmail.text               = "EMAIL ADDRESS"
    $titleEDEmail.AutoSize           = $true
    $titleEDEmail.width              = 457
    $titleEDEmail.height             = 142
    $titleEDEmail.Anchor             = 'top,right,left'
    $titleEDEmail.location           = New-Object System.Drawing.Point(5,130)
    $titleEDEmail.Font               = New-Object System.Drawing.Font('Consolas',15,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
    $titleEDEmail.ForeColor          = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

    $textEDEmail                     = New-Object System.Windows.Forms.TextBox
    $textEDEmail.width               = 480
    $textEDEmail.height              = 30
    $textEDEmail.Anchor              = 'top,right,left'
    $textEDEmail.location            = New-Object System.Drawing.Point(10,160)
    $textEDEmail.Font                = New-Object System.Drawing.Font('Consolas',9)
    $textEDEmail.ForeColor           = [System.Drawing.ColorTranslator]::FromHtml("#000000")
    $textEDEmail.ReadOnly            = $true

    $buttonEDUCheck                  = New-Object System.Windows.Forms.Button
    $buttonEDUCheck.FlatStyle        = 'Flat'
    $buttonEDUCheck.Text             = "Check AD Account"
    $buttonEDUCheck.width            = 230
    $buttonEDUCheck.height           = 30
    $buttonEDUCheck.Location         = New-Object System.Drawing.Point(260, 35)
    $buttonEDUCheck.Font             = New-Object System.Drawing.Font('Consolas',9)
    $buttonEDUCheck.ForeColor        = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

    $buttonEDDisable                 = New-Object System.Windows.Forms.Button
    $buttonEDDisable.FlatStyle       = 'Flat'
    $buttonEDDisable.Text            = "Disable Account"
    $buttonEDDisable.width           = 100
    $buttonEDDisable.height          = 60
    $buttonEDDisable.Location        = New-Object System.Drawing.Point(260, 230)
    $buttonEDDisable.Font            = New-Object System.Drawing.Font('Consolas',9)
    $buttonEDDisable.ForeColor       = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

    $buttonEDEnable                  = New-Object System.Windows.Forms.Button
    $buttonEDEnable.FlatStyle        = 'Flat'
    $buttonEDEnable.Text             = "Enable Account"
    $buttonEDEnable.width            = 100
    $buttonEDEnable.height           = 60
    $buttonEDEnable.Location         = New-Object System.Drawing.Point(390, 230)
    $buttonEDEnable.Font             = New-Object System.Drawing.Font('Consolas',9)
    $buttonEDEnable.ForeColor        = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

    $textEDAcctStatus                = New-Object System.Windows.Forms.Button
    $textEDAcctStatus.Text           = ""
    $textEDAcctStatus.width          = 230
    $textEDAcctStatus.height         = 60
    $textEDAcctStatus.Location       = New-Object System.Drawing.Point(10, 230)
    $textEDAcctStatus.Font           = New-Object System.Drawing.Font('Consolas',9)
    $textEDAcctStatus.ForeColor      = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
    $textEDAcctStatus.Enabled        = $false
    
    $titleEDAccStatus                = New-Object system.Windows.Forms.Label
    $titleEDAccStatus.text          = "ACCOUNT STATUS"
    $titleEDAccStatus.AutoSize      = $true
    $titleEDAccStatus.width         = 457
    $titleEDAccStatus.height        = 142
    $titleEDAccStatus.Anchor        = 'top,right,left'
    $titleEDAccStatus.location      = New-Object System.Drawing.Point(10,200)
    $titleEDAccStatus.Font          = New-Object System.Drawing.Font('Consolas',15,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
    $titleEDAccStatus.ForeColor     = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")


    $formEDAccount.controls.AddRange(@($titleEDAccountFName,$titleEDAccountFName,$titleEDAccountLName,$textEDAccountFirstN,$textEDAccountLastN,$titleEDUsername,$textEDUsername,$titleEDEmail,$textEDEmail,$textEDAcctStatus,$buttonEDEnable,$buttonEDUCheck,$buttonEDDisable,$titleEDAccStatus,$panelEDAccount))
    
    $varEDTime                     = Get-Date
    #Check AD Account
    $buttonEDUCheck.Add_Click( {
        #Text Field Variables
        $textEDAcctStatus.Clear
        $EDUNamePull                   = Get-ADUser -Identity $textEDUsername.text -Properties ObjectGUID, Name, LastLogonDate, LockedOut, AccountLockOutTime, Enabled, GivenName, Surname, UserPrincipalName
        $textEDAccountFirstN.text      = $EDUNamePull.GivenName
        $textEDAccountLastN.text       = $EDUNamePull.Surname
        $textEDEmail.text              = $EDUNamePull.UserPrincipalName
        $textEDUStatus                 = $EDUNamePull.Enabled


            if($textEDUStatus -eq $false){
                $textEDAcctStatus.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#b30000")
                $textEDAcctStatus.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
                $textEDAcctStatus.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
                $textEDAcctStatus.text = "Account is Disabled!"

            }
            else{
                $textEDAcctStatus.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#228c22")
                $textEDAcctStatus.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
                $textEDAcctStatus.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
                $textEDAcctStatus.text = "Account is Enabled!"
            }
        })
   

    #Disable Account
    $buttonEDDisable.Add_Click( {
        $textEDAcctStatus.Clear
        $EDUNamePull                = Get-ADUser -Identity $textEDUsername.text -Properties ObjectGUID, Name, LastLogonDate, LockedOut, AccountLockOutTime, Enabled, GivenName, Surname, UserPrincipalName, sAMAccountName
        $textEDAccountFirstN.text   = $EDUNamePull.GivenName
        $textEDAccountLastN.text    = $EDUNamePull.Surname
        $textEDEmail.text           = $EDUNamePull.UserPrincipalName
        $textEDUStatus              = $EDUNamePull.Enabled
   
        

        #Disable the Account via ObjectGUID
        Disable-ADAccount -Identity $EDUNamePull.sAMAccountName
        $EDResult = "Disable attempt for $($textEDEmail.text) was done at"
        Start-Sleep 3
        if($textEDUStatus -eq $false){
            $textEDAcctStatus.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#b30000")
            $textEDAcctStatus.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
            $textEDAcctStatus.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
            $textEDAcctStatus.text = "Account is Disabled!"

        }
        else{
            $textEDAcctStatus.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#228c22")
            $textEDAcctStatus.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
            $textEDAcctStatus.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
            $textEDAcctStatus.text = "Account is Enabled!"
        }
            #Log & report externally
        
            If (Test-Path $ReceiptsFolder\DisableEnableTool) {
                Write-Host "Logging changes..."
                
           }
            Else {
                Start-Sleep 1
                New-Item -Path "${ReceiptsFolder}\DisableEnableTool" -ItemType Directory
                Out-File -FilePath $ReceiptsFolder\DisableEnableTool\DisableEnableLog.txt -Encoding utf8
                Write-Host "Logging changes in a new folder..."
           }
            Add-Content -Path $ReceiptsFolder\DisableEnableTool\DisableEnableLog.txt "$EDResult $varEDTime"
       })
   

    $buttonEDEnable.Add_Click( {
        $textEDAcctStatus.Clear
        $EDUNamePull                = Get-ADUser -Identity $textEDUsername.text -Properties ObjectGUID, Name, LastLogonDate, LockedOut, AccountLockOutTime, Enabled, GivenName, Surname, UserPrincipalName
        $textEDAccountFirstN.text   = $EDUNamePull.GivenName
        $textEDAccountLastN.text    = $EDUNamePull.Surname
        $textEDEmail.text           = $EDUNamePull.UserPrincipalName
        $textEDUStatus              = $EDUNamePull.Enabled
   
        

        #Unlock the Account via ObjectGUID
        Enable-ADAccount -Identity $EDUNamePull.sAMAccountName
        $EDResult = "Enable attempt for $($textEDEmail.text) was done at"
        Start-Sleep 3
        if($textEDUStatus -eq $false){
            $textEDAcctStatus.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#b30000")
            $textEDAcctStatus.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
            $textEDAcctStatus.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
            $textEDAcctStatus.text = "Account is Disabled!"

        }
        else{
            $textEDAcctStatus.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#228c22")
            $textEDAcctStatus.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
            $textEDAcctStatus.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
            $textEDAcctStatus.text = "Account is Enabled!"
        }
            #Log & report externally
        
            If (Test-Path $ReceiptsFolder\DisableEnableTool) {
                Write-Host "Logging changes..."
                
           }
            Else {
                Start-Sleep 1
                New-Item -Path "${ReceiptsFolder}\DisableEnableTool" -ItemType Directory
                Out-File -FilePath $ReceiptsFolder\DisableEnableTool\DisableEnableLog.txt -Encoding utf8
                Write-Host "Logging changes in a new folder..."
           }
            Add-Content -Path $ReceiptsFolder\DisableEnableTool\DisableEnableLog.txt "$EDResult $varEDTime"
       })

    [void]$formEDAccount.ShowDialog()
    })


#Group Membership Tool 
$formGroupMembMang.Add_Click( { 
    $formGMembMgmt                   = New-Object System.Windows.Forms.Form
    $formGMembMgmt.ClientSize        = New-Object System.Drawing.Point(580,570)
    $formGMembMgmt.StartPosition     = 'CenterScreen'
    $formGMembMgmt.FormBorderStyle   = 'FixedSingle'
    $formGMembMgmt.MinimizeBox       = $false
    $formGMembMgmt.MaximizeBox       = $false
    $formGMembMgmt.ShowIcon          = $false
    $formGMembMgmt.Text              = "Group Membership Manager"
    $formGMembMgmt.TopMost           = $false
    $formGMembMgmt.AutoScroll        = $false
    $formGMembMgmt.BackColor         = [System.Drawing.ColorTranslator]::FromHtml("#252525")

    #Panel Setup
    $panelGrpMembMgmt                = New-Object System.Windows.Forms.Panel
    $panelGrpMembMgmt.height         = 510
    $panelGrpMembMgmt.width          = 572
    $panelGrpMembMgmt.Anchor         = 'top,right,left'
    $panelGrpMembMgmt.location       = New-Object System.Drawing.Point(10,10)
    $panelGrpMembMgmt.BackColor      = [System.Drawing.ColorTranslator]::FromHtml("#252525")

    #AD Unlock Tool
    $titleGMMgmtFirstN               = New-Object system.Windows.Forms.Label
    $titleGMMgmtFirstN.text          = "FIRST NAME"
    $titleGMMgmtFirstN.AutoSize      = $true
    $titleGMMgmtFirstN.width         = 457
    $titleGMMgmtFirstN.height        = 142
    $titleGMMgmtFirstN.Anchor        = 'top,right,left'
    $titleGMMgmtFirstN.location      = New-Object System.Drawing.Point(5,70)
    $titleGMMgmtFirstN.Font          = New-Object System.Drawing.Font('Consolas',15,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
    $titleGMMgmtFirstN.ForeColor     = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

    $titleGMMgmtLastN                = New-Object system.Windows.Forms.Label
    $titleGMMgmtLastN.text           = "LAST NAME"
    $titleGMMgmtLastN.AutoSize       = $true
    $titleGMMgmtLastN.width          = 457
    $titleGMMgmtLastN.height         = 142
    $titleGMMgmtLastN.Anchor         = 'top,right,left'
    $titleGMMgmtLastN.location       = New-Object System.Drawing.Point(255,70)
    $titleGMMgmtLastN.Font           = New-Object System.Drawing.Font('Consolas',15,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
    $titleGMMgmtLastN.ForeColor      = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

    $textGMMgmtFirstN                = New-Object System.Windows.Forms.TextBox
    $textGMMgmtFirstN.width          = 230
    $textGMMgmtFirstN.height         = 30
    $textGMMgmtFirstN.Anchor         = 'top,right,left'
    $textGMMgmtFirstN.location       = New-Object System.Drawing.Point(10,100)
    $textGMMgmtFirstN.Font           = New-Object System.Drawing.Font('Consolas',9)
    $textGMMgmtFirstN.ForeColor      = [System.Drawing.ColorTranslator]::FromHtml("#000000")
    $textGMMgmtFirstN.ReadOnly       = $true
    $textGMMgmtFirstN.text           = $null

    $textGMMgmtLastN                 = New-Object System.Windows.Forms.TextBox
    $textGMMgmtLastN.width           = 230
    $textGMMgmtLastN.height          = 30
    $textGMMgmtLastN.Anchor          = 'top,right,left'
    $textGMMgmtLastN.location        = New-Object System.Drawing.Point(260,100)
    $textGMMgmtLastN.Font            = New-Object System.Drawing.Font('Consolas',9)
    $textGMMgmtLastN.ForeColor       = [System.Drawing.ColorTranslator]::FromHtml("#000000")
    $textGMMgmtLastN.ReadOnly        = $true

    $titleGMMgmtUsern                = New-Object system.Windows.Forms.Label
    $titleGMMgmtUsern.text           = "USER NAME"
    $titleGMMgmtUsern.AutoSize       = $true
    $titleGMMgmtUsern.width          = 457
    $titleGMMgmtUsern.height         = 142
    $titleGMMgmtUsern.Anchor         = 'top,right,left'
    $titleGMMgmtUsern.location       = New-Object System.Drawing.Point(10,10)
    $titleGMMgmtUsern.Font           = New-Object System.Drawing.Font('Consolas',15,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
    $titleGMMgmtUsern.ForeColor      = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

    $textGMMgmtUserN                 = New-Object System.Windows.Forms.TextBox
    $textGMMgmtUserN.width           = 230
    $textGMMgmtUserN.height          = 30
    $textGMMgmtUserN.Anchor          = 'top,right,left'
    $textGMMgmtUserN.location        = New-Object System.Drawing.Point(10,40)
    $textGMMgmtUserN.Font            = New-Object System.Drawing.Font('Consolas',9)
    $textGMMgmtUserN.ForeColor       = [System.Drawing.ColorTranslator]::FromHtml("#000000")

    $titleGMMgmtEmail                = New-Object system.Windows.Forms.Label
    $titleGMMgmtEmail.text           = "EMAIL ADDRESS"
    $titleGMMgmtEmail.AutoSize       = $true
    $titleGMMgmtEmail.width          = 457
    $titleGMMgmtEmail.height         = 142
    $titleGMMgmtEmail.Anchor         = 'top,right,left'
    $titleGMMgmtEmail.location       = New-Object System.Drawing.Point(5,130)
    $titleGMMgmtEmail.Font           = New-Object System.Drawing.Font('Consolas',15,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
    $titleGMMgmtEmail.ForeColor      = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

    $textGmMgmtEmail                 = New-Object System.Windows.Forms.TextBox
    $textGmMgmtEmail.width           = 480
    $textGmMgmtEmail.height          = 30
    $textGmMgmtEmail.Anchor          = 'top,right,left'
    $textGmMgmtEmail.location        = New-Object System.Drawing.Point(10,160)
    $textGmMgmtEmail.Font            = New-Object System.Drawing.Font('Consolas',9)
    $textGmMgmtEmail.ForeColor       = [System.Drawing.ColorTranslator]::FromHtml("#000000")
    $textGmMgmtEmail.ReadOnly        = $true

    $buttonGMMgmtADChk              = New-Object System.Windows.Forms.Button
    $buttonGMMgmtADChk.FlatStyle    = 'Flat'
    $buttonGMMgmtADChk.Text         = "Check AD Account"
    $buttonGMMgmtADChk.width        = 230
    $buttonGMMgmtADChk.height       = 30
    $buttonGMMgmtADChk.Location     = New-Object System.Drawing.Point(260, 35)
    $buttonGMMgmtADChk.Font         = New-Object System.Drawing.Font('Consolas',9)
    $buttonGMMgmtADChk.ForeColor    = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
    
    $textGMMgmtGroup                = New-Object system.Windows.Forms.Label
    $textGMMgmtGroup.text           = "GROUP MANAGER"
    $textGMMgmtGroup.AutoSize       = $true
    $textGMMgmtGroup.width          = 457
    $textGMMgmtGroup.height         = 142
    $textGMMgmtGroup.Anchor         = 'top,right,left'
    $textGMMgmtGroup.location       = New-Object System.Drawing.Point(5,200)
    $textGMMgmtGroup.Font           = New-Object System.Drawing.Font('Consolas',15,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
    $textGMMgmtGroup.ForeColor      = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

    $textGMMgmtCurnt                = New-Object system.Windows.Forms.Label
    $textGMMgmtCurnt.text           = "ASSIGNED GROUPS"
    $textGMMgmtCurnt.AutoSize       = $true
    $textGMMgmtCurnt.width          = 457
    $textGMMgmtCurnt.height         = 142
    $textGMMgmtCurnt.Anchor         = 'top,right,left'
    $textGMMgmtCurnt.location       = New-Object System.Drawing.Point(65,230)
    $textGMMgmtCurnt.Font           = New-Object System.Drawing.Font('Consolas',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
    $textGMMgmtCurnt.ForeColor      = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
    
    
    $textGMMgmtAvail                = New-Object system.Windows.Forms.Label
    $textGMMgmtAvail.text           = "AVAILABLE GROUPS"
    $textGMMgmtAvail.AutoSize       = $true
    $textGMMgmtAvail.width          = 457
    $textGMMgmtAvail.height         = 142
    $textGMMgmtAvail.Anchor         = 'top,right,left'
    $textGMMgmtAvail.location       = New-Object System.Drawing.Point(365,230)
    $textGMMgmtAvail.Font           = New-Object System.Drawing.Font('Consolas',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
    $textGMMgmtAvail.ForeColor      = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

    $listGMMgmtAvail                = New-Object system.Windows.Forms.ListBox
    $listGMMgmtAvail.width          = 200
    $listGMMgmtAvail.height         = 300
    $listGMMgmtAvail.Anchor         = 'top,right,left'
    $listGMMgmtAvail.location       = New-Object System.Drawing.Point(330,250)
    $listGMMgmtAvail.Font           = New-Object System.Drawing.Font('Consolas',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
    $listGMMgmtAvail.ForeColor      = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
    $listGMMgmtAvail.BackColor      = [System.Drawing.ColorTranslator]::FromHtml("#252525")
    $listGMMgmtAvail.HorizontalScrollbar = $true
    $listGMMgmtAvail.ScrollAlwaysVisible = $true
    $listGMMgmtAvail.Sorted         = $true

    $listGMMgmtCurnt                = New-Object system.Windows.Forms.ListBox
    $listGMMgmtCurnt.width          = 200
    $listGMMgmtCurnt.height         = 300
    $listGMMgmtCurnt.Anchor         = 'top,right,left'
    $listGMMgmtCurnt.location       = New-Object System.Drawing.Point(30,250)
    $listGMMgmtCurnt.Font           = New-Object System.Drawing.Font('Consolas',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
    $listGMMgmtCurnt.ForeColor      = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
    $listGMMgmtCurnt.BackColor      = [System.Drawing.ColorTranslator]::FromHtml("#252525")
    $listGMMgmtCurnt.HorizontalScrollbar = $true
    $listGMMgmtCurnt.ScrollAlwaysVisible = $true
    $listGMMgmtCurnt.Sorted         = $true

    $groupRemoveUser                = New-Object System.Windows.Forms.Button
    $groupRemoveUser.Text           = ">>"
    $groupRemoveUser.width          = 30
    $groupRemoveUser.height         = 25
    $groupRemoveUser.Location       = New-Object System.Drawing.Point(265, 290)
    $groupRemoveUser.Font           = New-Object System.Drawing.Font('Consolas',9)
    $groupRemoveUser.ForeColor      = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

    $groupAddUser                   = New-Object System.Windows.Forms.Button
    $groupAddUser.Text              = "<<"
    $groupAddUser.width             = 30
    $groupAddUser.height            = 25
    $groupAddUser.Location          = New-Object System.Drawing.Point(265, 370)
    $groupAddUser.Font              = New-Object System.Drawing.Font('Consolas',9)
    $groupAddUser.ForeColor         = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

    $formGMembMgmt.controls.AddRange(@($textGMMgmtUserN,$textGMMgmtFirstN,$textGMMgmtLastN,$textGmMgmtEmail,$textGMMgmtGroup,$listGMMgmtCurnt,$listGMMgmtAvail,$titleGMMgmtFirstN,$titleGMMgmtFirstN,$titleGMMgmtLastN,$titleGMMgmtUsern,$titleGMMgmtEmail,$buttonGMMgmtADChk,$textGMMgmtAvail,$textGMMgmtCurnt,$groupAddUser,$groupRemoveUser,$panelGrpMembMgmt))    
    
    $AddingGroups = $listGMMgmtAvail.SelectedItems
    $RemovingGroups = $listGMMgmtCurnt.SelectedItems

    function Get-AllAdGroups {
        $allgroups = (Get-ADGroup -Filter *).DistinguishedName | ForEach-Object {
            ($_.Split(',') | Where-Object { $_.Contains("CN=") -and -not $_.Contains('Builtin') }).Replace("CN=", "")
        }
        [array]::Sort($allgroups)
        return $allgroups
        }

    function Get-AvailableADGroups {
        param (
            $Groups
        )
        
        $availableGroups = @()
    
        foreach ($g in $(Get-AllAdGroups)) {
            if ($Groups -notcontains $g) {
              $availableGroups += $g
            }
          }
        [array]::Sort($availableGroups)
    
        return $availableGroups
        
    }
    function Get-Difference {
        param (
            $Old, $New
        )
    
        $diff = @{}
        $compare = Compare-Object -ReferenceObject $Old -DifferenceObject $New
    
        foreach ($obj in $compare) {
            $diff[$obj.InputObject] = $obj.SideIndicator
        }
        return $diff
    }
    
    function Update-ADGroups {
        param ( $Groups, $NewGroups, $UserName )
        $diff = Get-Difference $Groups $NewGroups
        foreach ($obj in $diff.Keys) {
            if ($diff[$obj] -eq '<=') { Remove-ADGroupMember -Identity $obj -Members $UserName -Confirm:$false }
            if ($diff[$obj] -eq '=>') { Add-ADGroupMember -Identity $obj -Members $UserName -Confirm:$false }
        }
    }
    
    #region UI Stuff
    
    function Set-AvaliableADGroupsBox {
        param( $CurrentGroups )
        
        if (-not $CurrentGroups) {
            return
        }
        
        $availableGroups = Get-AvailableADGroups $availableGroups
        
        foreach ($group in $availableGroups) {
            $listGMMgmtAvail.Items.Add($group)
        }

    }
    
    function Set-CurrentADGroupsBox {
        param ( $CurrentGroups )
        
        if ( -not $CurrentGroups) { 
            return 
        }
        
        [array]::Sort($CurrentGroups)
        foreach ($group in $CurrentGroups) {
            $listGMMgmtCurnt.Items.Add($group)
        }
    }

    function Add-Group {
        param ( $GroupName )
        $listGMMgmtAvail.Items.Remove($GroupName)
        $listGMMgmtCurnt.Items.Add($GroupName)
        
    }
    
    function Remove-Group {
        param ( $GroupName )
        $listGMMgmtCurnt.Items.Remove($GroupName)
        $listGMMgmtAvail.Items.Add($GroupName)
        
    }

    $buttonGMMgmtADChk.Add_Click( {

        $GroupPull                     = Get-ADUser -Identity $textGMMgmtUserN.text -Properties GivenName, Surname, UserPrincipalName, samAccountName
        $textGMMgmtFirstN.text         = $GroupPull.GivenName
        $textGMMgmtLastN.text          = $GroupPull.Surname
        $textGmMgmtEmail.text          = $GroupPull.UserPrincipalName

        Get-AllAdGroups $allgroups
        Write-Host " test $allgroups "
        Get-AvailableADGroups -groups $allcurrentgroups
        Write-Host " test2 $allcurrentgroups "
        Set-AvaliableADGroupsBox -CurrentGroups $availableGroups
        Write-Host " test3 $availableGroups "
        Set-CurrentADGroupsBox
        Write-Host " test4 $CurrentGroups "

    })

    $groupAddUser.add_Click({
        if ( -not $listGMMgmtAvail.SelectedItem) {
            return
        }

        Add-Group $listGMMgmtAvail.SelectedItem
    
    })

    $groupRemoveUser.add_Click({
        if ( -not $listGMMgmtCurnt.SelectedItem) {
            return
        }
        Remove-Group $listGMMgmtCurnt.SelectedItem
    })
    [void]$formGMembMgmt.ShowDialog()


    } )


#Remote Mailbox Tool (COMPLETED)
$formEnableRMBX.Add_Click( { 
    $formEnableRMB                   = New-Object System.Windows.Forms.Form
    $formEnableRMB.ClientSize        = New-Object System.Drawing.Point(580,350)
    $formEnableRMB.StartPosition     = 'CenterScreen'
    $formEnableRMB.FormBorderStyle   = 'FixedSingle'
    $formEnableRMB.MinimizeBox       = $false
    $formEnableRMB.MaximizeBox       = $false
    $formEnableRMB.ShowIcon          = $false
    $formEnableRMB.Text              = "Enable Remote Mailbox"
    $formEnableRMB.TopMost           = $false
    $formEnableRMB.AutoScroll        = $false
    $formEnableRMB.BackColor         = [System.Drawing.ColorTranslator]::FromHtml("#252525")

    #Panel Setup
    $panelRMBSetup                   = New-Object System.Windows.Forms.Panel
    $panelRMBSetup.height            = 350
    $panelRMBSetup.width             = 572
    $panelRMBSetup.Anchor            = 'top,right,left'
    $panelRMBSetup.location          = New-Object System.Drawing.Point(10,10)
    $panelRMBSetup.BackColor         = [System.Drawing.ColorTranslator]::FromHtml("#252525")

    #Enable Disable Panels
    $titleEnableRMBFName             = New-Object system.Windows.Forms.Label
    $titleEnableRMBFName.text        = "FIRST NAME"
    $titleEnableRMBFName.AutoSize    = $true
    $titleEnableRMBFName.width       = 457
    $titleEnableRMBFName.height      = 142
    $titleEnableRMBFName.Anchor      = 'top,right,left'
    $titleEnableRMBFName.location    = New-Object System.Drawing.Point(5,70)
    $titleEnableRMBFName.Font        = New-Object System.Drawing.Font('Consolas',15,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
    $titleEnableRMBFName.ForeColor   = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

    $titleEnableRMBLName             = New-Object system.Windows.Forms.Label
    $titleEnableRMBLName.text        = "LAST NAME"
    $titleEnableRMBLName.AutoSize    = $true
    $titleEnableRMBLName.width       = 457
    $titleEnableRMBLName.height      = 142
    $titleEnableRMBLName.Anchor      = 'top,right,left'
    $titleEnableRMBLName.location    = New-Object System.Drawing.Point(255,70)
    $titleEnableRMBLName.Font        = New-Object System.Drawing.Font('Consolas',15,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
    $titleEnableRMBLName.ForeColor   = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

    $textEnableRMBFName               = New-Object System.Windows.Forms.TextBox
    $textEnableRMBFName.width         = 230
    $textEnableRMBFName.height        = 30
    $textEnableRMBFName.Anchor        = 'top,right,left'
    $textEnableRMBFName.location      = New-Object System.Drawing.Point(10,100)
    $textEnableRMBFName.Font          = New-Object System.Drawing.Font('Consolas',9)
    $textEnableRMBFName.ForeColor     = [System.Drawing.ColorTranslator]::FromHtml("#000000")
    $textEnableRMBFName.ReadOnly      = $true
    $textEnableRMBFName.text          = $null

    $textEnableRMBLName              = New-Object System.Windows.Forms.TextBox
    $textEnableRMBLName.width        = 230
    $textEnableRMBLName.height       = 30
    $textEnableRMBLName.Anchor       = 'top,right,left'
    $textEnableRMBLName.location     = New-Object System.Drawing.Point(260,100)
    $textEnableRMBLName.Font         = New-Object System.Drawing.Font('Consolas',9)
    $textEnableRMBLName.ForeColor    = [System.Drawing.ColorTranslator]::FromHtml("#000000")
    $textEnableRMBLName.ReadOnly     = $true

    $titleRMBUsername                = New-Object system.Windows.Forms.Label
    $titleRMBUsername.text           = "USER NAME"
    $titleRMBUsername.AutoSize       = $true
    $titleRMBUsername.width          = 457
    $titleRMBUsername.height         = 142
    $titleRMBUsername.Anchor         = 'top,right,left'
    $titleRMBUsername.location       = New-Object System.Drawing.Point(10,10)
    $titleRMBUsername.Font           = New-Object System.Drawing.Font('Consolas',15,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
    $titleRMBUsername.ForeColor      = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

    $textRMBUsername                 = New-Object System.Windows.Forms.TextBox
    $textRMBUsername.width           = 230
    $textRMBUsername.height          = 30
    $textRMBUsername.Anchor          = 'top,right,left'
    $textRMBUsername.location        = New-Object System.Drawing.Point(10,40)
    $textRMBUsername.Font            = New-Object System.Drawing.Font('Consolas',9)
    $textRMBUsername.ForeColor       = [System.Drawing.ColorTranslator]::FromHtml("#000000")

    $titleRMBEmail                   = New-Object system.Windows.Forms.Label
    $titleRMBEmail.text              = "EMAIL ADDRESS"
    $titleRMBEmail.AutoSize          = $true
    $titleRMBEmail.width             = 457
    $titleRMBEmail.height            = 142
    $titleRMBEmail.Anchor            = 'top,right,left'
    $titleRMBEmail.location          = New-Object System.Drawing.Point(5,130)
    $titleRMBEmail.Font              = New-Object System.Drawing.Font('Consolas',15,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
    $titleRMBEmail.ForeColor         = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

    $textRMBEmail                    = New-Object System.Windows.Forms.TextBox
    $textRMBEmail.width              = 480
    $textRMBEmail.height             = 30
    $textRMBEmail.Anchor             = 'top,right,left'
    $textRMBEmail.location           = New-Object System.Drawing.Point(10,160)
    $textRMBEmail.Font               = New-Object System.Drawing.Font('Consolas',9)
    $textRMBEmail.ForeColor          = [System.Drawing.ColorTranslator]::FromHtml("#000000")
    $textRMBEmail.ReadOnly           = $true

    $buttonRMBCheck                  = New-Object System.Windows.Forms.Button
    $buttonRMBCheck.FlatStyle        = 'Flat'
    $buttonRMBCheck.Text             = "Check AD Account"
    $buttonRMBCheck.width            = 230
    $buttonRMBCheck.height           = 30
    $buttonRMBCheck.Location         = New-Object System.Drawing.Point(260, 35)
    $buttonRMBCheck.Font             = New-Object System.Drawing.Font('Consolas',9)
    $buttonRMBCheck.ForeColor        = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

    $buttonRMBEnable                = New-Object System.Windows.Forms.Button
    $buttonRMBEnable.FlatStyle      = 'Flat'
    $buttonRMBEnable.Text           = "ENABLE REMOTE MAILBOX"
    $buttonRMBEnable.width          = 100
    $buttonRMBEnable.height         = 60
    $buttonRMBEnable.Location       = New-Object System.Drawing.Point(260, 230)
    $buttonRMBEnable.Font           = New-Object System.Drawing.Font('Consolas',9)
    $buttonRMBEnable.ForeColor      = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

    $textRMBAcctStatus               = New-Object System.Windows.Forms.Button
    $textRMBAcctStatus.Text          = ""
    $textRMBAcctStatus.width         = 230
    $textRMBAcctStatus.height        = 60
    $textRMBAcctStatus.Location      = New-Object System.Drawing.Point(10, 230)
    $textRMBAcctStatus.Font          = New-Object System.Drawing.Font('Consolas',9)
    $textRMBAcctStatus.ForeColor     = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
    $textRMBAcctStatus.Enabled       = $false
    
    $titleRMBAcctStatus              = New-Object system.Windows.Forms.Label
    $titleRMBAcctStatus.text         = "REMOTE MAILBOX STATUS"
    $titleRMBAcctStatus.AutoSize     = $true
    $titleRMBAcctStatus.width        = 457
    $titleRMBAcctStatus.height       = 142
    $titleRMBAcctStatus.Anchor       = 'top,right,left'
    $titleRMBAcctStatus.location     = New-Object System.Drawing.Point(10,200)
    $titleRMBAcctStatus.Font         = New-Object System.Drawing.Font('Consolas',15,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
    $titleRMBAcctStatus.ForeColor    = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")


    $formEnableRMB.controls.AddRange(@($titleEnableRMBFName,$titleEnableRMBFName,$titleEnableRMBLName,$textEnableRMBFName,$textEnableRMBLName,$titleRMBUsername,$textRMBUsername,$titleRMBEmail,$textRMBEmail,$textRMBAcctStatus,$buttonRMBCheck,$buttonRMBEnable,$titleRMBAcctStatus,$panelRMBSetup))
    
    $varRMBTime                      = Get-Date
    #Check for Remote Mailbox
    $buttonRMBCheck.Add_Click( {
        #Text Field Variables
        $textRMBAcctStatus.Clear
        $RMBNamePull.Clear
        $RMBNamePull                    = Get-ADUser -Identity $textRMBUsername.text -Properties  Name, GivenName, Surname, UserPrincipalName
        $textEnableRMBFName.text        = $RMBNamePull.GivenName
        $textEnableRMBLName.text        = $RMBNamePull.Surname
        $textRMBEmail.text              = $RMBNamePull.UserPrincipalName
       
        if(Get-RemoteMailbox -Identity $textRMBUsername.text -ea 0){
            $textRMBAcctStatus.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#228c22")
            $textRMBAcctStatus.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#ffffff")
            $textRMBAcctStatus.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
            $textRMBAcctStatus.text = "Account has a Remote Mailbox!"}
        else{
            $textRMBAcctStatus.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#b30000")
            $textRMBAcctStatus.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#ffffff")
            $textRMBAcctStatus.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
            $textRMBAcctStatus.text = "Account does not have a Remote Mailbox"
        }
         })
   

    #Enable Remote Mailbox
    $buttonRMBEnable.Add_Click( {
        #Text Field Variables
        $textRMBAcctStatus.Clear
        $RMBNamePull.Clear
        $RMBNamePull                    = Get-ADUser -Identity $textRMBUsername.text -Properties Name, GivenName, Surname, UserPrincipalName
        $textEnableRMBFName.text        = $RMBNamePull.GivenName
        $textEnableRMBLName.text        = $RMBNamePull.Surname
        $textRMBEmail.text              = $RMBNamePull.UserPrincipalName
       
        if(Get-RemoteMailbox -Identity $textRMBUsername.text -ea 0){
            $textRMBAcctStatus.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#228c22")
            $textRMBAcctStatus.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#ffffff")
            $textRMBAcctStatus.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
            $textRMBAcctStatus.text = "Account has a Remote Mailbox already!"}
        else{
            Enable-RemoteMailbox -Identity $textRMBUsername.text -RemoteRoutingAddress $textRMBEmail.text
            $textRMBAcctStatus.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#ffff00")
            $textRMBAcctStatus.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
            $textRMBAcctStatus.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
            $textRMBAcctStatus.text = "Setting Remote Mailbox... One moment"
            Start-Sleep -Seconds 2
            $textRMBAcctStatus.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#228c22")
            $textRMBAcctStatus.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#ffffff")
            $textRMBAcctStatus.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
            $textRMBAcctStatus.text = "Remote Mailbox set! Send an SMTP test ASAP!"
            $RMBResults = "Remote Mailbox set for $($RMBNamePull.UserPrincipalName) at: "
        }
            #Log & report externally
        
            If (Test-Path $ReceiptsFolder\RemoteMailboxTool) {
                Write-Host "Logging changes..."
                
           }
            Else {
                Start-Sleep 1
                New-Item -Path "${ReceiptsFolder}\RemoteMailboxTool" -ItemType Directory
                Out-File -FilePath $ReceiptsFolder\RemoteMailboxTool\RemoteMailboxLog.txt -Encoding utf8
                Write-Host "Logging changes in a new folder..."
           }
            Add-Content -Path $ReceiptsFolder\RemoteMailboxTool\RemoteMailboxLog.txt "$RMBResults $varRMBTime"
       }) 
   

    [void]$formEnableRMB.ShowDialog()
    })

#Rename Account Tool (COMPLETED)
$formRenameAcc.Add_Click( { 
    $formRNAccount                   = New-Object System.Windows.Forms.Form
    $formRNAccount.ClientSize        = New-Object System.Drawing.Point(580,600)
    $formRNAccount.StartPosition     = 'CenterScreen'
    $formRNAccount.FormBorderStyle   = 'FixedSingle'
    $formRNAccount.MinimizeBox       = $false
    $formRNAccount.MaximizeBox       = $false
    $formRNAccount.ShowIcon          = $false
    $formRNAccount.Text              = "Rename Account"
    $formRNAccount.TopMost           = $false
    $formRNAccount.AutoScroll        = $false
    $formRNAccount.BackColor         = [System.Drawing.ColorTranslator]::FromHtml("#252525")

    #Panel Setup
    $panelRNAccount                  = New-Object System.Windows.Forms.Panel
    $panelRNAccount.height           = 600
    $panelRNAccount.width            = 572
    $panelRNAccount.Anchor           = 'top,right,left'
    $panelRNAccount.location         = New-Object System.Drawing.Point(10,10)
    $panelRNAccount.BackColor        = [System.Drawing.ColorTranslator]::FromHtml("#252525")

    #Rename Panels
    $titleRNAccountFName             = New-Object system.Windows.Forms.Label
    $titleRNAccountFName.text        = "FIRST NAME"
    $titleRNAccountFName.AutoSize    = $true
    $titleRNAccountFName.width       = 457
    $titleRNAccountFName.height      = 142
    $titleRNAccountFName.Anchor      = 'top,right,left'
    $titleRNAccountFName.location    = New-Object System.Drawing.Point(5,70)
    $titleRNAccountFName.Font        = New-Object System.Drawing.Font('Consolas',15,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
    $titleRNAccountFName.ForeColor   = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

    $titleRNAccountLName             = New-Object system.Windows.Forms.Label
    $titleRNAccountLName.text        = "LAST NAME"
    $titleRNAccountLName.AutoSize    = $true
    $titleRNAccountLName.width       = 457
    $titleRNAccountLName.height      = 142
    $titleRNAccountLName.Anchor      = 'top,right,left'
    $titleRNAccountLName.location    = New-Object System.Drawing.Point(255,70)
    $titleRNAccountLName.Font        = New-Object System.Drawing.Font('Consolas',15,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
    $titleRNAccountLName.ForeColor   = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

    $textRNAccountFirstN             = New-Object System.Windows.Forms.TextBox
    $textRNAccountFirstN.width       = 230
    $textRNAccountFirstN.height      = 30
    $textRNAccountFirstN.Anchor      = 'top,right,left'
    $textRNAccountFirstN.location    = New-Object System.Drawing.Point(10,100)
    $textRNAccountFirstN.Font        = New-Object System.Drawing.Font('Consolas',9)
    $textRNAccountFirstN.ForeColor   = [System.Drawing.ColorTranslator]::FromHtml("#000000")
    $textRNAccountFirstN.ReadOnly    = $true
    $textRNAccountFirstN.text        = $null

    $textRNAccountLastN              = New-Object System.Windows.Forms.TextBox
    $textRNAccountLastN.width        = 230
    $textRNAccountLastN.height       = 30
    $textRNAccountLastN.Anchor       = 'top,right,left'
    $textRNAccountLastN.location     = New-Object System.Drawing.Point(260,100)
    $textRNAccountLastN.Font         = New-Object System.Drawing.Font('Consolas',9)
    $textRNAccountLastN.ForeColor    = [System.Drawing.ColorTranslator]::FromHtml("#000000")
    $textRNAccountLastN.ReadOnly     = $true

    $titleRNUsername                 = New-Object system.Windows.Forms.Label
    $titleRNUsername.text            = "USER NAME"
    $titleRNUsername.AutoSize        = $true
    $titleRNUsername.width           = 457
    $titleRNUsername.height          = 142
    $titleRNUsername.Anchor          = 'top,right,left'
    $titleRNUsername.location        = New-Object System.Drawing.Point(10,10)
    $titleRNUsername.Font            = New-Object System.Drawing.Font('Consolas',15,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
    $titleRNUsername.ForeColor       = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

    $textRNUsername                  = New-Object System.Windows.Forms.TextBox
    $textRNUsername.width            = 230
    $textRNUsername.height           = 30
    $textRNUsername.Anchor           = 'top,right,left'
    $textRNUsername.location         = New-Object System.Drawing.Point(10,40)
    $textRNUsername.Font             = New-Object System.Drawing.Font('Consolas',9)
    $textRNUsername.ForeColor        = [System.Drawing.ColorTranslator]::FromHtml("#000000")

    $titleRNEmail                    = New-Object system.Windows.Forms.Label
    $titleRNEmail.text               = "EMAIL ADDRESS"
    $titleRNEmail.AutoSize           = $true
    $titleRNEmail.width              = 457
    $titleRNEmail.height             = 142
    $titleRNEmail.Anchor             = 'top,right,left'
    $titleRNEmail.location           = New-Object System.Drawing.Point(5,130)
    $titleRNEmail.Font               = New-Object System.Drawing.Font('Consolas',15,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
    $titleRNEmail.ForeColor          = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

    $textRNEmail                     = New-Object System.Windows.Forms.TextBox
    $textRNEmail.width               = 480
    $textRNEmail.height              = 30
    $textRNEmail.Anchor              = 'top,right,left'
    $textRNEmail.location            = New-Object System.Drawing.Point(10,160)
    $textRNEmail.Font                = New-Object System.Drawing.Font('Consolas',9)
    $textRNEmail.ForeColor           = [System.Drawing.ColorTranslator]::FromHtml("#000000")
    $textRNEmail.ReadOnly            = $true

    $buttonRNUCheck                  = New-Object System.Windows.Forms.Button
    $buttonRNUCheck.FlatStyle        = 'Flat'
    $buttonRNUCheck.Text             = "Check AD Account"
    $buttonRNUCheck.width            = 230
    $buttonRNUCheck.height           = 30
    $buttonRNUCheck.Location         = New-Object System.Drawing.Point(260, 35)
    $buttonRNUCheck.Font             = New-Object System.Drawing.Font('Consolas',9)
    $buttonRNUCheck.ForeColor        = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

    $buttonRNChangeName              = New-Object System.Windows.Forms.Button
    $buttonRNChangeName.FlatStyle    = 'Flat'
    $buttonRNChangeName.Text         = "Rename Account"
    $buttonRNChangeName.width        = 560
    $buttonRNChangeName.height       = 30
    $buttonRNChangeName.Location     = New-Object System.Drawing.Point(10, 435)
    $buttonRNChangeName.Font         = New-Object System.Drawing.Font('Consolas',9)
    $buttonRNChangeName.ForeColor    = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

    $textRNAcctLog                   = New-Object System.Windows.Forms.RichTextBox
    $textRNAcctLog.Text              = ""
    $textRNAcctLog.width             = 560
    $textRNAcctLog.height            = 100
    $textRNAcctLog.Location          = New-Object System.Drawing.Point(10, 475)
    $textRNAcctLog.Font              = New-Object System.Drawing.Font('Consolas',9)
    $textRNAcctLog.BackColor         = [System.Drawing.ColorTranslator]::FromHtml("#252525")
    $textRNAcctLog.ForeColor         = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
    $textRNAcctLog.ReadONly          = $true
    
    $titleRNAccStatus                = New-Object system.Windows.Forms.Label
    $titleRNAccStatus.text           = "NEW ACCOUNT INFO"
    $titleRNAccStatus.AutoSize       = $true
    $titleRNAccStatus.width          = 457
    $titleRNAccStatus.height         = 142
    $titleRNAccStatus.Anchor         = 'top,right,left'
    $titleRNAccStatus.location       = New-Object System.Drawing.Point(5,220)
    $titleRNAccStatus.Font           = New-Object System.Drawing.Font('Consolas',15,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
    $titleRNAccStatus.ForeColor      = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
#New Account Info
    $titleNewAccountFName            = New-Object system.Windows.Forms.Label
    $titleNewAccountFName.text       = "FIRST NAME"
    $titleNewAccountFName.AutoSize   = $true
    $titleNewAccountFName.width      = 457
    $titleNewAccountFName.height     = 142
    $titleNewAccountFName.Anchor     = 'top,right,left'
    $titleNewAccountFName.location   = New-Object System.Drawing.Point(5,310)
    $titleNewAccountFName.Font       = New-Object System.Drawing.Font('Consolas',15,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
    $titleNewAccountFName.ForeColor  = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

    $titleNewAccountLName            = New-Object system.Windows.Forms.Label
    $titleNewAccountLName.text       = "LAST NAME"
    $titleNewAccountLName.AutoSize   = $true
    $titleNewAccountLName.width      = 457
    $titleNewAccountLName.height     = 142
    $titleNewAccountLName.Anchor     = 'top,right,left'
    $titleNewAccountLName.location   = New-Object System.Drawing.Point(255,310)
    $titleNewAccountLName.Font       = New-Object System.Drawing.Font('Consolas',15,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
    $titleNewAccountLName.ForeColor  = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

    $textNewAccountFirstN            = New-Object System.Windows.Forms.TextBox
    $textNewAccountFirstN.width      = 230
    $textNewAccountFirstN.height     = 30
    $textNewAccountFirstN.Anchor     = 'top,right,left'
    $textNewAccountFirstN.location   = New-Object System.Drawing.Point(10,340)
    $textNewAccountFirstN.Font       = New-Object System.Drawing.Font('Consolas',9)
    $textNewAccountFirstN.ForeColor  = [System.Drawing.ColorTranslator]::FromHtml("#000000")
    $textNewAccountFirstN.text       = $null

    $textNewAccountLastN             = New-Object System.Windows.Forms.TextBox
    $textNewAccountLastN.width       = 230
    $textNewAccountLastN.height      = 30
    $textNewAccountLastN.Anchor      = 'top,right,left'
    $textNewAccountLastN.location    = New-Object System.Drawing.Point(260,340)
    $textNewAccountLastN.Font        = New-Object System.Drawing.Font('Consolas',9)
    $textNewAccountLastN.ForeColor   = [System.Drawing.ColorTranslator]::FromHtml("#000000")

    $titleNewUsername                = New-Object system.Windows.Forms.Label
    $titleNewUsername.text           = "USER NAME"
    $titleNewUsername.AutoSize       = $true
    $titleNewUsername.width          = 457
    $titleNewUsername.height         = 142
    $titleNewUsername.Anchor         = 'top,right,left'
    $titleNewUsername.location       = New-Object System.Drawing.Point(5,250)
    $titleNewUsername.Font           = New-Object System.Drawing.Font('Consolas',15,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
    $titleNewUsername.ForeColor      = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

    $textNewUsername                 = New-Object System.Windows.Forms.TextBox
    $textNewUsername.width           = 230
    $textNewUsername.height          = 30
    $textNewUsername.Anchor          = 'top,right,left'
    $textNewUsername.location        = New-Object System.Drawing.Point(10,280)
    $textNewUsername.Font            = New-Object System.Drawing.Font('Consolas',9)
    $textNewUsername.ForeColor       = [System.Drawing.ColorTranslator]::FromHtml("#000000")

    $titleNewDomainN                 = New-Object system.Windows.Forms.Label
    $titleNewDomainN.text            = "NEW DOMAIN NAME"
    $titleNewDomainN.AutoSize        = $true
    $titleNewDomainN.width           = 457
    $titleNewDomainN.height          = 142
    $titleNewDomainN.Anchor          = 'top,right,left'
    $titleNewDomainN.location        = New-Object System.Drawing.Point(260,250)
    $titleNewDomainN.Font            = New-Object System.Drawing.Font('Consolas',15,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
    $titleNewDomainN.ForeColor       = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

    $textNewDomainN                  = New-Object System.Windows.Forms.ComboBox
    $textNewDomainN.width            = 230
    $textNewDomainN.height           = 30
    $textNewDomainN.Anchor           = 'top,right,left'
    $textNewDomainN.location         = New-Object System.Drawing.Point(260,280)
    $textNewDomainN.Font             = New-Object System.Drawing.Font('Consolas',9)
    $textNewDomainN.ForeColor        = [System.Drawing.ColorTranslator]::FromHtml("#000000")
    (Get-ADForest).UPNSuffixes | ForEach-Object {[void] $textNewDomainN.Items.Add($_)}

    $titleNewEmail                   = New-Object system.Windows.Forms.Label
    $titleNewEmail.text              = "EMAIL ADDRESS"
    $titleNewEmail.AutoSize          = $true
    $titleNewEmail.width             = 457
    $titleNewEmail.height            = 142
    $titleNewEmail.Anchor            = 'top,right,left'
    $titleNewEmail.location          = New-Object System.Drawing.Point(5,370)
    $titleNewEmail.Font              = New-Object System.Drawing.Font('Consolas',15,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
    $titleNewEmail.ForeColor         = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

    $textNewEmail                    = New-Object System.Windows.Forms.TextBox
    $textNewEmail.width              = 480
    $textNewEmail.height             = 30
    $textNewEmail.Anchor             = 'top,right,left'
    $textNewEmail.location           = New-Object System.Drawing.Point(10,400)
    $textNewEmail.Font               = New-Object System.Drawing.Font('Consolas',9)
    $textNewEmail.ForeColor          = [System.Drawing.ColorTranslator]::FromHtml("#000000")
   
    $formRNAccount.controls.AddRange(@($titleRNAccountFName,$titleRNAccountFName,$titleRNAccountLName,$textRNAccountFirstN,$textRNAccountLastN,$titleRNUsername,$textRNUsername,$titleRNEmail,$textRNEmail,$textRNAcctLog,$buttonRNUCheck,$buttonRNChangeName,$titleRNAccStatus,$titleNewAccountFName,$titleNewAccountLName,$textNewAccountFirstN,$textNewAccountLastN,$titleNewUsername,$textNewUsername,$titleNewEmail,$textNewEmail,$textNewDomainN,$titleNewDomainN,$panelRNAccount))
    
    $varEDTime                       = Get-Date

    #Check AD Account
    $buttonRNUCheck.Add_Click( {
        #Text Field Variables
        $textRNAcctLog.Clear
        $RNUNamePull                   = Get-ADUser -Identity $textRNUsername.text -Properties ObjectGUID, Name, LastLogonDate, LockedOut, AccountLockOutTime, Enabled, GivenName, Surname, UserPrincipalName, sAMAccountName
        $textRNAccountFirstN.text      = $RNUNamePull.GivenName
        $textRNAccountLastN.text       = $RNUNamePull.Surname
        $textRNEmail.text              = $RNUNamePull.UserPrincipalName

        #Set New Account Info
        $textNewUsername.text           = $RNUNamePull.sAMAccountName
        $textNewAccountFirstN.text      = $RNUNamePull.GivenName
        $textNewAccountLastN.text       = $RNUNamePull.Surname
        $textNewEmail.text              = $RNUNamePull.UserPrincipalName
        })
    
    #Account Name Updater
    $textNewDomainN.add_textchanged({$textNewEmail.Text = "$($textNewUsername.text)@$($textNewDomainN.text)"})
    $textNewUsername.add_textchanged({$textNewEmail.Text = "$($textNewUsername.text)@$($textNewDomainN.text)"})


    #Change Account Name
    $buttonRNChangeName.Add_Click( {
        $textRNAcctLog.Clear        
        #Check if new user is taken
        $varNewUser                    = $($textNewUsername.text)
        $varNewRNCheck                 = Get-ADUser -filter {sAMAccountName -eq $varNewUser}
        $varRenameLock                 = Get-ADUser -Identity $textRNUsername.text -Properties ObjectGUID

            if ($varNewRNCheck -ne $null) {
                $varRenameLog = "The account $($textNewUsername.text) already exists. Please try a new name."
                $textRNAcctLog.text = $varRenameLog
            }
            else{
                if(Get-RemoteMailbox -Identity $textNewUsername.text -ea 0){
                    Set-RemoteMailbox -Identity $textNewUsername.text -PrimarySMTPAddress $textNewEmail.text
                    Start-Sleep -Seconds 1
                    Set-RemoteMailbox -Identity $textNewUsername.text -RemoteRoutingAddress $textNewEmail.text -EmailAddresses "smtp:$($textRNEmail.text)"
                    Start-Sleep -Seconds 1
                    Set-ADUser -Identity $varRenameLock.ObjectGUID -GivenName $textNewAccountFirstN.text -Surname $textNewAccountLastN.text -SamAccountName $textNewUsername.text -UserPrincipalName $textNewEmail.text -DisplayName "$($extNewAccountFirstN.text) $($textNewAccountLastN.text)"
                    Start-Sleep -Seconds 1
              <#    Set-ADUser -Identity $textNewUsername.text -Surname $varNewHALastN
                    Start-Sleep -Seconds 1
                    Set-ADUser -Identity $textNewUsername.text -DisplayName "$varNewHAFirstN $varNewHALastN"
                    Start-Sleep -Seconds 1#>
                    $varRenameLog = "The account $($textNewUsername.text) has been renamed. `nPlease check in AD if the account changes took and repair as necessary.`n New details:`nUsername: $($textNewUsername.text)`nUPN: $($textNewEmail.text)`nName: $($textNewAccountFirstN.text) $($textNewAccountLastN.text)`nDisplay Name:$($extNewAccountFirstN.text) $($textNewAccountLastN.text)"
                    $textRNAcctLog.text = $varRenameLog

                    $RNResult = "Rename attempt for $($textRNEmail.text) was done at"
                    #Log & report externally
                }
                else{
                    Set-ADUser -Identity $varRenameLock.ObjectGUID -GivenName $textNewAccountFirstN.text -Surname $textNewAccountLastN.text -SamAccountName $textNewUsername.text -UserPrincipalName $textNewEmail.text -DisplayName "$($extNewAccountFirstN.text) $($textNewAccountLastN.text)"
                    Start-Sleep -Seconds 1
          <#        Set-ADUser -Identity $textNewUsername.text -Surname $varNewHALastN
                    Start-Sleep -Seconds 1
                    Set-ADUser -Identity $textNewUsername.text -DisplayName "$varNewHAFirstN $varNewHALastN"
                    Start-Sleep -Seconds 1#>
                    $varRenameLog = "The account $($textNewUsername.text) has been renamed. `nPlease check in AD if the account changes took and repair as necessary.`n New details:`nUsername: $($textNewUsername.text)`nUPN: $($textNewEmail.text)`nName: $($textNewAccountFirstN.text) $($textNewAccountLastN.text)`nDisplay Name:$($extNewAccountFirstN.text) $($textNewAccountLastN.text)"
                    $textRNAcctLog.text = $varRenameLog

                    $RNResult = "Rename attempt for $($textRNEmail.text) was done at"

                }
        
                If (Test-Path $ReceiptsFolder\RenameTool) {
                    Write-Host "Logging changes..."
                
                }
                 Else {
                    Start-Sleep 1
                    New-Item -Path "${ReceiptsFolder}\RenameTool" -ItemType Directory
                    Out-File -FilePath $ReceiptsFolder\RenameTool\RenameLog.txt -Encoding utf8
                    Write-Host "Logging changes in a new folder..."
            }
            Add-Content -Path $ReceiptsFolder\RenameTool\RenameLog.txt "$RNResult $varEDTime"
        }
    })

    [void]$formRNAccount.ShowDialog()
    })


#Unlock Account Tool (COMPLETED)
$formUnlockAcc.Add_Click( { 
    $formADUnlock                    = New-Object System.Windows.Forms.Form
    $formADUnlock.ClientSize         = New-Object System.Drawing.Point(580,480)
    $formADUnlock.StartPosition      = 'CenterScreen'
    $formADUnlock.FormBorderStyle    = 'FixedSingle'
    $formADUnlock.MinimizeBox        = $false
    $formADUnlock.MaximizeBox        = $false
    $formADUnlock.ShowIcon           = $false
    $formADUnlock.Text               = "Unlock AD Account"
    $formADUnlock.TopMost            = $false
    $formADUnlock.AutoScroll         = $false
    $formADUnlock.BackColor          = [System.Drawing.ColorTranslator]::FromHtml("#252525")

    #Panel Setup
    $panelADUnlocker                 = New-Object System.Windows.Forms.Panel
    $panelADUnlocker.height          = 350
    $panelADUnlocker.width           = 572
    $panelADUnlocker.Anchor          = 'top,right,left'
    $panelADUnlocker.location        = New-Object System.Drawing.Point(10,10)
    $panelADUnlocker.BackColor       = [System.Drawing.ColorTranslator]::FromHtml("#252525")

    #AD Unlock Tool
    $titleADUFirstName               = New-Object system.Windows.Forms.Label
    $titleADUFirstName.text          = "FIRST NAME"
    $titleADUFirstName.AutoSize      = $true
    $titleADUFirstName.width         = 457
    $titleADUFirstName.height        = 142
    $titleADUFirstName.Anchor        = 'top,right,left'
    $titleADUFirstName.location      = New-Object System.Drawing.Point(5,70)
    $titleADUFirstName.Font          = New-Object System.Drawing.Font('Consolas',15,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
    $titleADUFirstName.ForeColor     = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

    $titleADULastName                = New-Object system.Windows.Forms.Label
    $titleADULastName.text           = "LAST NAME"
    $titleADULastName.AutoSize       = $true
    $titleADULastName.width          = 457
    $titleADULastName.height         = 142
    $titleADULastName.Anchor         = 'top,right,left'
    $titleADULastName.location       = New-Object System.Drawing.Point(255,70)
    $titleADULastName.Font           = New-Object System.Drawing.Font('Consolas',15,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
    $titleADULastName.ForeColor      = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

    $textADUFirstName                = New-Object System.Windows.Forms.TextBox
    $textADUFirstName.width          = 230
    $textADUFirstName.height         = 30
    $textADUFirstName.Anchor         = 'top,right,left'
    $textADUFirstName.location       = New-Object System.Drawing.Point(10,100)
    $textADUFirstName.Font           = New-Object System.Drawing.Font('Consolas',9)
    $textADUFirstName.ForeColor      = [System.Drawing.ColorTranslator]::FromHtml("#000000")
    $textADUFirstName.ReadOnly       = $true
    $textADUFirstName.text           = $null

    $textADULastName                 = New-Object System.Windows.Forms.TextBox
    $textADULastName.width           = 230
    $textADULastName.height          = 30
    $textADULastName.Anchor          = 'top,right,left'
    $textADULastName.location        = New-Object System.Drawing.Point(260,100)
    $textADULastName.Font            = New-Object System.Drawing.Font('Consolas',9)
    $textADULastName.ForeColor       = [System.Drawing.ColorTranslator]::FromHtml("#000000")
    $textADULastName.ReadOnly        = $true

    $titleADUUsername                = New-Object system.Windows.Forms.Label
    $titleADUUsername.text           = "USER NAME"
    $titleADUUsername.AutoSize       = $true
    $titleADUUsername.width          = 457
    $titleADUUsername.height         = 142
    $titleADUUsername.Anchor         = 'top,right,left'
    $titleADUUsername.location       = New-Object System.Drawing.Point(10,10)
    $titleADUUsername.Font           = New-Object System.Drawing.Font('Consolas',15,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
    $titleADUUsername.ForeColor      = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

    $textADUUsername                 = New-Object System.Windows.Forms.TextBox
    $textADUUsername.width           = 230
    $textADUUsername.height          = 30
    $textADUUsername.Anchor          = 'top,right,left'
    $textADUUsername.location        = New-Object System.Drawing.Point(10,40)
    $textADUUsername.Font            = New-Object System.Drawing.Font('Consolas',9)
    $textADUUsername.ForeColor       = [System.Drawing.ColorTranslator]::FromHtml("#000000")

    $titleADUEmail                   = New-Object system.Windows.Forms.Label
    $titleADUEmail.text              = "EMAIL ADDRESS"
    $titleADUEmail.AutoSize          = $true
    $titleADUEmail.width             = 457
    $titleADUEmail.height            = 142
    $titleADUEmail.Anchor            = 'top,right,left'
    $titleADUEmail.location          = New-Object System.Drawing.Point(5,130)
    $titleADUEmail.Font              = New-Object System.Drawing.Font('Consolas',15,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
    $titleADUEmail.ForeColor         = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

    $textADUEmail                    = New-Object System.Windows.Forms.TextBox
    $textADUEmail.width              = 480
    $textADUEmail.height             = 30
    $textADUEmail.Anchor             = 'top,right,left'
    $textADUEmail.location           = New-Object System.Drawing.Point(10,160)
    $textADUEmail.Font               = New-Object System.Drawing.Font('Consolas',9)
    $textADUEmail.ForeColor          = [System.Drawing.ColorTranslator]::FromHtml("#000000")
    $textADUEmail.ReadOnly           = $true

    $buttonADUCheck                  = New-Object System.Windows.Forms.Button
    $buttonADUCheck.FlatStyle        = 'Flat'
    $buttonADUCheck.Text             = "Check AD Account"
    $buttonADUCheck.width            = 230
    $buttonADUCheck.height           = 30
    $buttonADUCheck.Location         = New-Object System.Drawing.Point(260, 35)
    $buttonADUCheck.Font             = New-Object System.Drawing.Font('Consolas',9)
    $buttonADUCheck.ForeColor        = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

    $buttonADUUnlock                 = New-Object System.Windows.Forms.Button
    $buttonADUUnlock.FlatStyle       = 'Flat'
    $buttonADUUnlock.Text            = "Unlock AD Account"
    $buttonADUUnlock.width           = 100
    $buttonADUUnlock.height          = 60
    $buttonADUUnlock.Location        = New-Object System.Drawing.Point(260, 230)
    $buttonADUUnlock.Font            = New-Object System.Drawing.Font('Consolas',9)
    $buttonADUUnlock.ForeColor       = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

    $buttonADPWReset                 = New-Object System.Windows.Forms.Button
    $buttonADPWReset.FlatStyle       = 'Flat'
    $buttonADPWReset.Text            = "Set New PW"
    $buttonADPWReset.width           = 100
    $buttonADPWReset.height          = 30
    $buttonADPWReset.Location        = New-Object System.Drawing.Point(260, 330)
    $buttonADPWReset.Font            = New-Object System.Drawing.Font('Consolas',9)
    $buttonADPWReset.ForeColor       = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

    $textADUAccStatus                = New-Object System.Windows.Forms.Button
    $textADUAccStatus.Text           = ""
    $textADUAccStatus.width          = 230
    $textADUAccStatus.height         = 60
    $textADUAccStatus.Location       = New-Object System.Drawing.Point(10, 230)
    $textADUAccStatus.Font           = New-Object System.Drawing.Font('Consolas',9)
    $textADUAccStatus.ForeColor      = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
    $textADUAccStatus.Enabled        = $false

    $textPWRAccStatus                = New-Object System.Windows.Forms.Button
    $textPWRAccStatus.Text           = ""
    $textPWRAccStatus.width          = 230
    $textPWRAccStatus.height         = 60
    $textPWRAccStatus.Location       = New-Object System.Drawing.Point(10, 360)
    $textPWRAccStatus.Font           = New-Object System.Drawing.Font('Consolas',9)
    $textPWRAccStatus.ForeColor      = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
    $textPWRAccStatus.Enabled        = $false

    $textPWRNewPassword              = New-Object System.Windows.Forms.MaskedTextBox
    $textPWRNewPassword.width        = 230
    $textPWRNewPassword.height       = 30
    $textPWRNewPassword.Anchor       = 'top,right,left'
    $textPWRNewPassword.location     = New-Object System.Drawing.Point(10,330)
    $textPWRNewPassword.Font         = New-Object System.Drawing.Font('Consolas',9)
    $textPWRNewPassword.ForeColor    = [System.Drawing.ColorTranslator]::FromHtml("#000000")
    $textPWRNewPassword.text         = $null
    $textPWRNewPassword.PasswordChar = "*"

    
    $titleADUAccStatus               = New-Object system.Windows.Forms.Label
    $titleADUAccStatus.text          = "ACCOUNT STATUS"
    $titleADUAccStatus.AutoSize      = $true
    $titleADUAccStatus.width         = 457
    $titleADUAccStatus.height        = 142
    $titleADUAccStatus.Anchor        = 'top,right,left'
    $titleADUAccStatus.location      = New-Object System.Drawing.Point(10,200)
    $titleADUAccStatus.Font          = New-Object System.Drawing.Font('Consolas',15,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
    $titleADUAccStatus.ForeColor     = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

    $titlePWResetTitle               = New-Object system.Windows.Forms.Label
    $titlePWResetTitle.text          = "RESET ACCOUNT PASSWORD"
    $titlePWResetTitle.AutoSize      = $true
    $titlePWResetTitle.width         = 457
    $titlePWResetTitle.height        = 142
    $titlePWResetTitle.Anchor        = 'top,right,left'
    $titlePWResetTitle.location      = New-Object System.Drawing.Point(10,300)
    $titlePWResetTitle.Font          = New-Object System.Drawing.Font('Consolas',15,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
    $titlePWResetTitle.ForeColor     = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")


    $formADUnlock.controls.AddRange(@($textPWRAccStatus,$titlePWResetTitle,$titleADUFirstName,$titleADUFirstName,$titleADULastName,$textADUFirstName,$textADULastName,$titleADUUsername,$textADUUsername,$titleADUEmail,$textADUEmail,$textADUAccStatus,$textPWRNewPassword,$buttonADPWReset,$buttonADUCheck,$buttonADUUnlock,$titleADUAccStatus,$panelADUnlocker))
    
    $varADPWRTime                    = Get-Date
    #Check AD Account
    $buttonADUCheck.Add_Click( {
            #Text Field Variables
            $textADUAccStatus.Clear
            $ADUNamePull                = Get-ADUser -Identity $textADUUsername.text -Properties ObjectGUID, Name, LastLogonDate, LockedOut, AccountLockOutTime, Enabled, GivenName, Surname, UserPrincipalName
            $textADUFirstName.text      = $ADUNamePull.GivenName
            $textADULastName.text       = $ADUNamePull.Surname
            $textADUEmail.text          = $ADUNamePull.UserPrincipalName
            $textADUStatus              = $ADUNamePull.LockedOut


            if($textADUStatus -eq $false){
                $textADUAccStatus.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#228c22")
                $textADUAccStatus.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#ffffff")
                $textADUAccStatus.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
                $textADUAccStatus.text = "Account is unlocked!"

            }
            else{
                $textADUAccStatus.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#b30000")
                $textADUAccStatus.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#ffffff")
                $textADUAccStatus.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
                $textADUAccStatus.text = "Account is locked!"
            }
        })
   

    #Unlock AD Account
    $buttonADUUnlock.Add_Click( {
        $textADUAccStatus.Clear
        $ADUNamePull                = Get-ADUser -Identity $textADUUsername.text -Properties ObjectGUID, Name, LastLogonDate, LockedOut, AccountLockOutTime, Enabled, GivenName, Surname, UserPrincipalName
        $textADUFirstName.text      = $ADUNamePull.GivenName
        $textADULastName.text       = $ADUNamePull.Surname
        $textADUEmail.text          = $ADUNamePull.UserPrincipalName
        $textADUStatus              = $ADUNamePull.LockedOut
        $ADUUnlockThis              = $ADUNamePull.ObjectGUID
        

        #Unlock the Account via ObjectGUID
        Unlock-ADAccount -Identity $ADUUnlockThis
        $ADUResult = "Unlock attempt for $textADUEmail.text was done at"
        
        if($textADUStatus -eq $false){
            $textADUAccStatus.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#228c22")
            $textADUAccStatus.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
            $textADUAccStatus.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
            $textADUAccStatus.text = "Account is unlocked!"

        }
        else{
            $textADUAccStatus.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#b30000")
            $textADUAccStatus.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
            $textADUAccStatus.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
            $textADUAccStatus.text = "Account is locked!"
        }
            #Log & report externally
        
            If (Test-Path $ReceiptsFolder\ADUnlockResetTool) {
                Write-Host "Logging changes..."
                
           }
            Else {
                Start-Sleep 1
                New-Item -Path "${ReceiptsFolder}\ADUnlockResetTool" -ItemType Directory
                Out-File -FilePath $ReceiptsFolder\ADUnlockResetTool\Unlocklog.txt -Encoding utf8
                Write-Host "Logging changes in a new folder..."
           }
            Add-Content -Path $ReceiptsFolder\ADUnlockResetTool\Unlocklog.txt "$ADUResult $varADPWRTime"
       })
   

    $buttonADPWReset.Add_Click( {
        $textPWRAccStatus.Clear
        $textPWRAccStatus.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
        $textPWRAccStatus.ForeColor = ""
        
        try{    
            if($textADUFirstName.text -gt 1 -and $textPWRNewPassword.text -gt 8){
                $varNewPassword   =  $textPWRNewPassword.text | ConvertTo-SecureString -AsPlainText -Force
                Set-ADAccountPassword -Identity $textADUUsername.text -NewPassword $varNewPassword -Reset
                Set-ADUser -Identity $textADUUsername.text -ChangePasswordAtLogon $true
                Unlock-ADAccount -Identity $textADUUsername.text

                $textPWRAccStatus.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#228c22")
                $textPWRAccStatus.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
                $textPWRAccStatus.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
                $textPWRAccStatus.text = "The password has been reset!"
                $ADRPWRResult = "Password attempt for $($textADUUsername.text) was successful on $varADPWRTime"
            }
            else{
                $textPWRAccStatus.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#b30000")
                $textPWRAccStatus.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
                $textPWRAccStatus.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
                $textPWRAccStatus.text = "Please select an account and a complex password! "
                $ADRPWRResult = "Password attempt for $($textADUUsername.text) has failed on $varADPWRTime"
            }

             }
        catch {   
                $textPWRAccStatus.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#b30000")
                $textPWRAccStatus.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
                $textPWRAccStatus.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
                $textPWRAccStatus.text = "Try a different password."
                $ADRPWRResult = "Password attempt for $($textADUUsername.text) has failed on $varADPWRTime"
        }
    

    #Log & report externally
        
         If (Test-Path $ReceiptsFolder\ADUnlockResetTool) {
             Write-Host "Logging changes..."
             
        }
         Else {
             Start-Sleep 1
             New-Item -Path "${ReceiptsFolder}\ADUnlockResetTool" -ItemType Directory
             Out-File -FilePath $ReceiptsFolder\ADUnlockResetTool\Resetlog.txt -Encoding utf8
             Write-Host "Logging changes in a new folder..."
        }
         Add-Content -Path $ReceiptsFolder\ADUnlockResetTool\Resetlog.txt "$ADRPWRResult + $varADPWRTime"
    })

    [void]$formADUnlock.ShowDialog()
    })


#RunADSync (COMPLETED)
$RunADSync.Add_Click( {
    $formADSync                      = New-Object System.Windows.Forms.Form
    $formADSync.ClientSize           = New-Object System.Drawing.Point(580,270)
    $formADSync.StartPosition        = 'CenterScreen'
    $formADSync.FormBorderStyle      = 'FixedSingle'
    $formADSync.MinimizeBox          = $false
    $formADSync.MaximizeBox          = $false
    $formADSync.ShowIcon             = $false
    $formADSync.Text                 = "Run AD Sync"
    $formADSync.TopMost              = $false
    $formADSync.AutoScroll           = $false
    $formADSync.BackColor            = [System.Drawing.ColorTranslator]::FromHtml("#252525")

#Panel Setup
    $panelRunADSync                  = New-Object System.Windows.Forms.Panel
    $panelRunADSync.height           = 250
    $panelRunADSync.width            = 572
    $panelRunADSync.Anchor           = 'top,right,left'
    $panelRunADSync.location         = New-Object System.Drawing.Point(10,10)
    $panelRunADSync.BackColor        = [System.Drawing.ColorTranslator]::FromHtml("#252525")

#Run AD Sync Panel

    $textADSyncLogs                  = New-Object System.Windows.Forms.RichTextBox
    $textADSyncLogs.width            = 560
    $textADSyncLogs.height           = 200
    $textADSyncLogs.location         = New-Object System.Drawing.Point(10,50)
    $textADSyncLogs.Font             = New-Object System.Drawing.Font('Consolas',9)
    $textADSyncLogs.BackColor        = [System.Drawing.ColorTranslator]::FromHtml("#252525")
    $textADSyncLogs.ForeColor        = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
    $textADSyncLogs.ReadOnly         = $true

    $buttonRunADSync                 = New-Object System.Windows.Forms.Button
    $buttonRunADSync.FlatStyle       = 'Flat'
    $buttonRunADSync.Text            = "RUN AD SYNC"
    $buttonRunADSync.width           = 560
    $buttonRunADSync.height          = 30
    $buttonRunADSync.Location        = New-Object System.Drawing.Point(10, 10)
    $buttonRunADSync.Font            = New-Object System.Drawing.Font('Consolas',9)
    $buttonRunADSync.ForeColor       = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")



#Hybrid Account Range Control
    $formADSync.controls.AddRange(@($buttonRunADSync,$textADSyncLogs,$panelRunADSync))

#Set up Button Clicks
    $buttonRunADSync.Add_Click( {
     
        $doSync = 1
    if ($doSync -eq 1) {
        $textADSyncLogs.AppendText("`nStarting Sync - Please Wait...")
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
                        $textADSyncLogs.AppendText("`nService ADSync is stuck. `nElevated permissions are required to restart the service... n`Attempting to restart")
                        Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass `"Restart-Service -Name ADSync -Force`"" -Verb RunAs -Wait
                        $textADSyncLogs.AppendText("`nADSync Service should be restarted. Trying DeltaSync again...")
                        $i = 2
                    } else {
                        if ($i -eq 1) {$textADSyncLogs.AppendText("`nAD Sync is already running. Attempting to try again...")}
                        if ($i % 5  -eq 0) {$textADSyncLogs.AppendText("`nStill trying...")}
                        Sleep 1
                        $i++
                        Continue
                    }
                }
            } while (!$result)
            Start-ADSyncSyncCycle -PolicyType Delta | Out-Null
        }
        $textADSyncLogs.AppendText("`nSuccess! Please allow a few minutes for changes to occur in the cloud.")
        Start-Sleep 1
}
Else {
    $textADSyncLogs.AppendText("`nError running AD Sync. Please try again manually.")
        exit    
}

    })
    #End Button Click
    [void]$formADSync.ShowDialog()
})


#End Script    
[void]$formMain.ShowDialog()

