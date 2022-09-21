<#Requires Administrator Rights

#Synopsis
This script adjusts existing FSLogix disk sizes. It may also resize User Profile Disks (VHDXs) because the FSLogix app has native hyper-v support.

To start, we need to set our variables:
1.) Target the app location for frx.exe (No trailing slashes)
2.) Target the path for the RDS Profiles (No trailing slashes)
3.) Target the disk size.

The script will then proceed to rename all VHDs in the subdirectories of your folder. (If you use VHDXs, change the script to VHDX)
Next, it will migrate the existing drive, to the new drive with the adjusted storage size.

You can verify the new storage quota by logging in to the account and opening the frxtray.exe app, and seeing the listed disk size.

Please be aware, this script was made by someone of a very novice level.
Test before pushing to production.
Use with caution.

- Jesse Campanella
#>

Read-Host "This script target all user profiles in a given folder, and will then resize them to the specified amount by renaming the current disk > cloning to a new disk. Press enter to continue...";


Read-Host "We're going to set some variables for the FSLogix App path, and the RDS Profile stores. Press enter to continue...";



$FSLogixAppPath = Read-Host -Prompt 'input the path for the FSLogix frx.exe app (Usually c:\Program Files\FSLogix\Apps)(DO NOT leave a trailing \)'
$RDSProfilePath = Read-Host -Prompt 'input the path for the FSLogix Profiles (DO NOT leave a trailing \)'
$NewDiskSize = Read-Host -Prompt "input the disk size in actual MBs, enter 5120 for 5gb, for example. "

Read-Host  "The profiles under $RDSProfilePath will be set to allow upto $NewDiskSize MB of space, using the frx.exe app found under $FSLogixAppPath. Press enter to continue..."


Read-Host 'Press enter to start...';


#Set the destination for the App
$ResizeCommand = $FSLogixAppPath+"\frx.exe"


#Rename ProfileDisks to Old
Get-ChildItem $RDSProfilePath -Directory | ForEach-Object {
    #Select the original VHD to be renamed Old
    $RenameSource = Join-Path $_.FullName -ChildPath *.vhd* -Resolve
    #Lock the Filepath of the original VHD
    $UserName = Join-Path $_.FullName -ChildPath *.vhd* -Resolve
    #Set Destination for creating new VHD
    $NewSource = $UserName -f $_.Name


    #Rename VHD to Old
    Rename-Item -Path $RenameSource -NewName "OldProfile.vhd"


    #Find Old VHD
    $OldSource = Join-Path $_.FullName -ChildPath *old* -Resolve


    #Target current Old VHD, set new file name to the previous Filepath, Set the DiskSize
    $size = "-size-mbs="+$NewDiskSize
    & $ResizeCommand migrate-vhd -src $OldSource -dest $NewSource $size -dynamic=1

}


