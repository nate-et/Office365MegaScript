# Notes: To come in future release:
# 
# 1. Better Definitions for "Type" of teams group
# 2. Fix the Teams stuff because it's broken currently
#
# ------------------------------------------------------------------------------------------------------------
#
#
# Change List
#
# (2.4.0)
# 1. Added Public Folder functionality
#
# (2.3.0)
# 1. Added Membership for Security Groups rather than just names
#
# (2.2.3)
# 1. Added debugging
#
# (2.2.0)
# 1. Added Security Groups Section
#
# (2.1.5)
# 1. Aletered wording on some of the verbose outputting
#
# (2.1.4)
# 1. Added Privacy to Teams Groups Sheet (e.g Public or Private)
# 2. Added confirmation prompt at end of export process to prompt if you want to open the exported file or not
#
#
#


# Set the path for the log file
$logFile = "C:\Temp\megascript_debug.txt"

# Start the transcript to capture all output and errors
Start-Transcript -Path $logFile -Append

# Install necessary modules if not already installed
if (-not (Get-Module -Name ImportExcel -ListAvailable)) {
    Install-Module -Name ImportExcel
}

if (-not (Get-Module -Name ExchangeOnlineManagement -ListAvailable)) {
    Install-Module -Name ExchangeOnlineManagement
}

if (-not (Get-Module -Name AzureAD -ListAvailable)) {
    Install-Module -Name AzureAD
}

if (-not (Get-Module -Name MSOnline -ListAvailable)) {
    Install-Module -Name MSOnline
}

# Import the module
Import-Module ImportExcel

Write-Host "Connecting To Exchange Online, AzureAD & MsolService..."
Connect-ExchangeOnline
Connect-AzureAD
Connect-MsolService

Write-Host "The output will be saved to your desktop"

Start-Sleep 2

# Prompt user for file name and path
$fileName = Read-Host -Prompt "Enter the file name (without extension):"
$filePath = [System.IO.Path]::Combine([System.Environment]::GetFolderPath("Desktop"), "$fileName.xlsx")

# Define worksheet names
$worksheets = @(
    "UserMailboxes_DelegatedAccess",
    "LicensedAccounts",
    "SharedMailboxes_DelegatedAccess",
    "DistributionLists_Members",
    "TeamsGroups_Members",
    "Security_Groups",
    "PublicFolderMailboxes",
    "PublicFolders"
)

# Create an Excel package
$excelPackage = New-Object OfficeOpenXml.ExcelPackage
$worksheets | ForEach-Object {
    $excelPackage.Workbook.Worksheets.Add($_) | Out-Null
}

# List of shared mailboxes with any mailbox access
Write-Host "Fetching Shared Mailboxes lists with members..."
$sharedMailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox
$totalSharedMailboxes = $sharedMailboxes.Count
$currentSharedMailbox = 0
$sharedMailboxesData = $sharedMailboxes | foreach {
    $currentSharedMailbox++
    $progressPercent = [math]::Round(($currentSharedMailbox / $totalSharedMailboxes) * 100)
    Write-Progress -Activity "Fetching Shared Mailboxes" -Status "Processing $($_.DisplayName)" -PercentComplete $progressPercent
    
    $mailbox = $_.DisplayName
    $accesses = @()
    
    # Fetch all mailbox permissions
    $permissions = Get-MailboxPermission -Identity $_.Identity -ErrorAction SilentlyContinue
    if ($permissions) {
        foreach ($permission in $permissions) {
            $accesses += [PSCustomObject]@{
                "SharedMailbox" = $mailbox
                "User" = $permission.User
                "AccessRights" = $permission.AccessRights -join ', '  # Convert array to comma-separated string
            }
        }
    } else {
        Write-Host "No mailbox access found for shared mailbox: $mailbox"
    }
    
    $accesses
}

# Export shared mailboxes with delegated mailbox access to Excel
$sharedMailboxesData | Export-Excel -ExcelPackage $excelPackage -WorksheetName $worksheets[2] -Title "Shared Mailboxes with Delegated Access" -BoldTopRow -PassThru | Out-Null

# List of distribution lists with members
Write-Host "Fetching Distribution Group lists with members..."
$distributionLists = Get-DistributionGroup
$totalDistributionLists = $distributionLists.Count
$currentDistributionList = 0
$distributionListsData = $distributionLists | foreach {
    $currentDistributionList++
    $progressPercent = [math]::Round(($currentDistributionList / $totalDistributionLists) * 100)
    Write-Progress -Activity "Fetching Distribution Lists" -Status "Processing $($_.DisplayName)" -PercentComplete $progressPercent
    
    $groupName = $_.DisplayName
    $members = Get-DistributionGroupMember -Identity $_.Identity -ErrorAction SilentlyContinue
    if ($members) {
        $members | foreach {
            [PSCustomObject]@{
                "DistributionList" = $groupName
                "MemberName" = $_.DisplayName
                "MemberEmail" = $_.PrimarySmtpAddress
            }
        }
    } else {
        [PSCustomObject]@{
            "DistributionList" = $groupName
            "MemberName" = "No members found"
            "MemberEmail" = "N/A"
        }
    }
}

# Export distribution lists with members to Excel
$distributionListsData | Export-Excel -ExcelPackage $excelPackage -WorksheetName $worksheets[3] -Title "Distribution Lists with Members" -BoldTopRow -PassThru | Out-Null

# List of all Teams and Groups with their members
Write-Host "Fetching Teams and Groups lists with members..."
$teamsGroups = @()

# Get all Microsoft 365 groups
$teamsAndGroups = Get-AzureADGroup -All $true | 
                    Where-Object { $_.GroupTypes -contains "Unified" } | 
                    Select-Object DisplayName, Description, ObjectId

$totalTeamsAndGroups = $teamsAndGroups.Count
$currentTeamOrGroup = 0

# Create an array to store teams and groups information
$teamsAndGroupsWithMembers = @()

# Iterate through each team/group
foreach ($group in $teamsAndGroups) {
    $currentTeamOrGroup++
    $progressPercent = [math]::Round(($currentTeamOrGroup / $totalTeamsAndGroups) * 100)
    Write-Progress -Activity "Fetching Teams and Groups" -Status "Processing $($group.DisplayName)" -PercentComplete $progressPercent
    
    # Fetch members of the current team/group
    $groupMembers = Get-AzureADGroupMember -All $true -ObjectId $group.ObjectId | 
                    Select-Object -ExpandProperty DisplayName

    # Create an object for each group member
    foreach ($member in $groupMembers) {
        $groupObject = [PSCustomObject]@{
            "GroupDisplayName" = $group.DisplayName
            "GroupDescription" = $group.Description
            "GroupMember" = $member
        }

        # Add the group object to the array
        $teamsAndGroupsWithMembers += $groupObject
    }
}

# Add a new worksheet for teams and groups
$teamsAndGroupsSheet = $excelPackage.Workbook.Worksheets["Teams_And_Groups"]
$teamsAndGroupsSheet.Cells.Clear()

# Add headers
$teamsAndGroupsSheet.Cells[1, 1].Value = "Group Display Name"
$teamsAndGroupsSheet.Cells[1, 2].Value = "Group Description"
$teamsAndGroupsSheet.Cells[1, 3].Value = "Group Member"

# Fill the sheet with the teams and groups data
$row = 2
foreach ($group in $teamsAndGroupsWithMembers) {
    $teamsAndGroupsSheet.Cells[$row, 1].Value = $group.GroupDisplayName
    $teamsAndGroupsSheet.Cells[$row, 2].Value = $group.GroupDescription
    $teamsAndGroupsSheet.Cells[$row, 3].Value = $group.GroupMember
    $row++
}

# Export Teams and Groups with members to Excel
$teamsGroups | Export-Excel -ExcelPackage $excelPackage -WorksheetName $worksheets[4] -Title "Teams and Groups with Members" -BoldTopRow -PassThru | Out-Null

Write-Host "Fetching Security Groups list with members..."

# Get all security groups with DisplayName and Description
$securityGroups = Get-AzureADGroup -All $true | 
                    Where-Object { $_.SecurityEnabled -eq $true } | 
                    Select-Object DisplayName, Description, ObjectId

$totalSecurityGroups = $securityGroups.Count
$currentSecurityGroup = 0

# Create an array to store security group information
$securityGroupsWithMembers = @()

# Iterate through each security group
foreach ($group in $securityGroups) {
    $currentSecurityGroup++
    $progressPercent = [math]::Round(($currentSecurityGroup / $totalSecurityGroups) * 100)
    Write-Progress -Activity "Fetching Security Groups" -Status "Processing $($group.DisplayName)" -PercentComplete $progressPercent
    
    # Fetch members of the current security group
    $groupMembers = Get-AzureADGroupMember -All $true -ObjectId $group.ObjectId | 
                    Select-Object -ExpandProperty DisplayName

    # Create an object for each group member
    foreach ($member in $groupMembers) {
        $groupObject = [PSCustomObject]@{
            "GroupDisplayName" = $group.DisplayName
            "GroupDescription" = $group.Description
            "GroupMember" = $member
        }

        # Add the group object to the array
        $securityGroupsWithMembers += $groupObject
    }
}

# Add a new worksheet for security groups
$securityGroupsSheet = $excelPackage.Workbook.Worksheets["Security_Groups"]
$securityGroupsSheet.Cells.Clear()

# Add headers
$securityGroupsSheet.Cells[1, 1].Value = "Group Display Name"
$securityGroupsSheet.Cells[1, 2].Value = "Group Member"
$securityGroupsSheet.Cells[1, 3].Value = "Group Description"

# Fill the sheet with the security group data
$row = 2
foreach ($group in $securityGroupsWithMembers) {
    $securityGroupsSheet.Cells[$row, 1].Value = $group.GroupDisplayName
    $securityGroupsSheet.Cells[$row, 2].Value = $group.GroupMember
    $securityGroupsSheet.Cells[$row, 3].Value = $group.GroupDescription
    $row++
}

# Export Security Groups to Excel
$securityGroupsWithMembers | Export-Excel -ExcelPackage $excelPackage -WorksheetName $worksheets[5] -Title "Security Groups with Members" -BoldTopRow -PassThru | Out-Null

Write-Host "Fetching Public Folder Mailboxes..."
$publicFolderMailboxes = Get-Mailbox -PublicFolder
$totalPublicFolderMailboxes = $publicFolderMailboxes.Count
$currentPublicFolderMailbox = 0

$publicFolderMailboxesData = $publicFolderMailboxes | foreach {
    $currentPublicFolderMailbox++
    $progressPercent = [math]::Round(($currentPublicFolderMailbox / $totalPublicFolderMailboxes) * 100)
    Write-Progress -Activity "Fetching Public Folder Mailboxes" -Status "Processing $($_.DisplayName)" -PercentComplete $progressPercent
    
    [PSCustomObject]@{
        "MailboxName" = $_.Name
        "PrimarySMTP" = $_.PrimarySmtpAddress
        "Alias" = $_.Alias
    }
}

# Export Public Folder Mailboxes to Excel
$publicFolderMailboxesData | Export-Excel -ExcelPackage $excelPackage -WorksheetName $worksheets[6] -Title "Public Folder Mailboxes" -BoldTopRow -PassThru | Out-Null

# Define function to get all public folders and their child folders
function Get-PublicFolderTree {
    param (
        [string]$ParentFolderPath = "\",  # Start from the root folder
        [int]$ParentId = 0,              # Parent ID for progress tracking
        [ref]$FolderIndex = 0            # Reference variable to track the folder index
    )
    
    # Get the public folders under the specified parent folder
    $publicFolders = Get-PublicFolder -Identity $ParentFolderPath -Recurse -ResultSize Unlimited | Select-Object Name, Identity, MailEnabled, ParentPath
    
    $folderData = @()

    $totalFolders = $publicFolders.Count

    foreach ($folder in $publicFolders) {
        # Increment the folder index
        $FolderIndex.Value++

        # Display progress
        $progressPercent = [math]::Round(($FolderIndex.Value / ($totalFolders + $ParentId)) * 100)
        Write-Progress -Activity "Fetching Public Folders" -Status "Processing $($folder.Identity)" -PercentComplete $progressPercent

        # Add folder information to the array
        $folderData += [PSCustomObject]@{
            "FolderName"     = $folder.Name
            "FolderPath"     = $folder.Identity
            "ParentPath"     = $folder.ParentPath
            "MailEnabled"    = if ($folder.MailEnabled) { "Yes" } else { "No" }
            "MailAddresses"  = if ($folder.MailEnabled) { ($folder.EmailAddresses | Where-Object { $_ -like 'SMTP:*' }) -join ', ' } else { "N/A" }
        }
        
        # Recursively get child folders if any
        $childFolders = Get-PublicFolderTree -ParentFolderPath $folder.Identity -ParentId $FolderIndex.Value -FolderIndex ([ref]$FolderIndex.Value)
        if ($childFolders) {
            $folderData += $childFolders
        }
    }
    
    return $folderData
}

Write-Host "Fetching Public Folders and their child folders recursively..."
$FolderIndex = 0
$publicFolders = Get-PublicFolderTree -FolderIndex ([ref]$FolderIndex)

# Check if a worksheet named "PublicFolders" already exists, and remove it if it does
$existingWorksheet = $excelPackage.Workbook.Worksheets["PublicFolders"]
if ($existingWorksheet) {
    $excelPackage.Workbook.Worksheets.Delete($existingWorksheet)
}

# Add a new worksheet for public folders
$publicFoldersSheet = $excelPackage.Workbook.Worksheets["PublicFolders"]

# Add headers
$publicFoldersSheet.Cells[1, 1].Value = "Folder Name"
$publicFoldersSheet.Cells[1, 2].Value = "Folder Path"
$publicFoldersSheet.Cells[1, 3].Value = "Parent Path"
$publicFoldersSheet.Cells[1, 4].Value = "Mail Enabled"
$publicFoldersSheet.Cells[1, 5].Value = "Mail Addresses"

# Fill the sheet with the public folders data
$row = 2
foreach ($folder in $publicFolders) {
    $publicFoldersSheet.Cells[$row, 1].Value = $folder.FolderName
    $publicFoldersSheet.Cells[$row, 2].Value = $folder.FolderPath
    $publicFoldersSheet.Cells[$row, 3].Value = $folder.ParentPath
    $publicFoldersSheet.Cells[$row, 4].Value = $folder.MailEnabled
    $publicFoldersSheet.Cells[$row, 5].Value = $folder.MailAddresses
    $row++
}

# Export public folders to Excel
$publicFolders | Export-Excel -ExcelPackage $excelPackage -WorksheetName $worksheets[7] -Title "Public Folders and Child Folders" -BoldTopRow -PassThru | Out-Null

# Save the Excel file
$excelPackage.SaveAs($filePath)

Write-Host "Export completed. The file has been saved to $filePath"

# Stop the transcript
Stop-Transcript

# Prompt the user to open the Excel file
$openFile = Read-Host "Do you want to open the Excel file now? (y/n)"
if ($openFile -eq "y") {
    Start-Process $filePath
}

