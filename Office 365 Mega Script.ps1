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

# Install ImportExcel module if not already installed
if (-not (Get-Module -Name ImportExcel -ListAvailable)) {
    Install-Module -Name ImportExcel
}

# Install ExchangeOnline module if not already installed
if (-not (Get-Module -Name ExchangeOnlineManagement -ListAvailable)) {
    Install-Module -Name ExchangeOnlineManagement
}

# Install AzureAD module if not already installed
if (-not (Get-Module -Name AzureAD -ListAvailable)) {
    Install-Module -Name AzureAD
}

# Install MSOnline module if not already installed
if (-not (Get-Module -Name MSOnline -ListAvailable)) {
    Install-Module -Name MSOnline
}

# Import the module
Import-Module ImportExcel

# Prompt for Exchange Online credentials
#$credential = Get-Credential -Message "Enter your Exchange Online admin credentials"


Write-Host "Connecting To Exchange Online, AzureAD & MsolService..."
# Connect to Required Modules
Connect-ExchangeOnline #-Credential $credential
Connect-AzureAD #-Credential $credential
Connect-MsolService #-Credential $credential

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
    "Security_Groups"
)


# Create an Excel package
$excelPackage = New-Object OfficeOpenXml.ExcelPackage
$worksheets | ForEach-Object {
    $excelPackage.Workbook.Worksheets.Add($_) | Out-Null
}


# List of user mailboxes with any mailbox access
Write-Host "Fetching User Mailboxes with Delegated Access list..."
$userMailboxes = Get-Mailbox | foreach {
    $mailbox = $_.DisplayName
    $accesses = @()
    
    # Fetch all mailbox permissions
    $permissions = Get-MailboxPermission -Identity $_.Identity -ErrorAction SilentlyContinue
    if ($permissions) {
        foreach ($permission in $permissions) {
            $accesses += [PSCustomObject]@{
                "Mailbox" = $mailbox
                "User" = $permission.User
                "AccessRights" = $permission.AccessRights -join ', '  # Convert array to comma-separated string
            }
        }
    } else {
        Write-Host "No mailbox access found for mailbox: $mailbox"
    }
    
    $accesses
}

# Export user mailboxes with delegated mailbox access to Excel
$userMailboxes | Export-Excel -ExcelPackage $excelPackage -WorksheetName $worksheets[0] -Title "User Mailboxes with Delegated Access" -BoldTopRow -PassThru | Out-Null



# List of accounts with assigned licenses
Write-Host "Fetching Licensed Users list..."
$licensedAccounts = Get-MsolUser -All | where {$_.isLicensed -eq "True"} | foreach {
    [PSCustomObject]@{
        "UPN" = $_.UserPrincipalName
        "License" = ($_.Licenses | ForEach-Object { $_.AccountSkuId }) -join ', '  # Convert array to comma-separated string
    }
}

# Export licensed accounts to Excel
$licensedAccounts | Export-Excel -ExcelPackage $excelPackage -WorksheetName $worksheets[1] -Title "Licensed Accounts" -BoldTopRow -PassThru  | Out-Null

# List of shared mailboxes with any mailbox access
Write-Host "Fetching Shared Mailboxes lists with members..."
$sharedMailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox | foreach {
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
$sharedMailboxes | Export-Excel -ExcelPackage $excelPackage -WorksheetName $worksheets[2] -Title "Shared Mailboxes with Delegated Access" -BoldTopRow -PassThru  | Out-Null

# List of public folder mailboxes with their aliases and email addresses
Write-Host "Fetching Public Folder Mailboxes list with aliases and email addresses..."
$publicFolderMailboxes = Get-Mailbox -PublicFolder | foreach {
    [PSCustomObject]@{
        "MailboxName" = $_.DisplayName
        "Alias" = $_.Alias
        "EmailAddresses" = ($_.EmailAddresses | Where-Object { $_ -like 'SMTP:*' }) -join ', '  # Filter SMTP addresses and join them into a single string
    }
}

# Check if a worksheet named "PublicFolderMailboxes" already exists, and remove it if it does
$existingWorksheet = $excelPackage.Workbook.Worksheets["PublicFolderMailboxes"]
if ($existingWorksheet) {
    $excelPackage.Workbook.Worksheets.Delete($existingWorksheet)
}

# Add a new worksheet for public folder mailboxes
$publicFolderSheet = $excelPackage.Workbook.Worksheets.Add("PublicFolderMailboxes")

# Add headers
$publicFolderSheet.Cells[1, 1].Value = "Mailbox Name"
$publicFolderSheet.Cells[1, 2].Value = "Alias"
$publicFolderSheet.Cells[1, 3].Value = "Email Addresses"

# Fill the sheet with the public folder mailboxes data
$row = 2
foreach ($mailbox in $publicFolderMailboxes) {
    $publicFolderSheet.Cells[$row, 1].Value = $mailbox.MailboxName
    $publicFolderSheet.Cells[$row, 2].Value = $mailbox.Alias
    $publicFolderSheet.Cells[$row, 3].Value = $mailbox.EmailAddresses
    $row++
}

# Export public folder mailboxes to Excel
$publicFolderMailboxes | Export-Excel -ExcelPackage $excelPackage -WorksheetName "PublicFolderMailboxes" -Title "Public Folder Mailboxes with Aliases and Email Addresses" -BoldTopRow -PassThru | Out-Null

# List of public folder mailboxes with their aliases and email addresses
Write-Host "Fetching Public Folder Mailboxes list with aliases and email addresses..."
# [Existing code]

# Recursive function to get all public folders and their child folders
function Get-PublicFolderTree {
    param (
        [string]$ParentFolderPath = "\\"
    )
    
    # Get all child public folders for the current parent path
    $publicFolders = Get-PublicFolder -Identity $ParentFolderPath -Recurse | Select-Object Name, Identity, MailEnabled, ParentPath
    
    $folderData = @()
    
    foreach ($folder in $publicFolders) {
        # Add folder information to the array
        $folderData += [PSCustomObject]@{
            "FolderName"     = $folder.Name
            "FolderPath"     = $folder.Identity
            "ParentPath"     = $folder.ParentPath
            "MailEnabled"    = if ($folder.MailEnabled) { "Yes" } else { "No" }
            "MailAddresses"  = if ($folder.MailEnabled) { ($folder.EmailAddresses | Where-Object { $_ -like 'SMTP:*' }) -join ', ' } else { "N/A" }
        }
        
        # Recursively get child folders
        $folderData += Get-PublicFolderTree -ParentFolderPath $folder.Identity
    }
    
    return $folderData
}

Write-Host "Fetching Public Folders and their child folders recursively..."
$publicFolders = Get-PublicFolderTree

# Check if a worksheet named "PublicFolders" already exists, and remove it if it does
$existingWorksheet = $excelPackage.Workbook.Worksheets["PublicFolders"]
if ($existingWorksheet) {
    $excelPackage.Workbook.Worksheets.Delete($existingWorksheet)
}

# Add a new worksheet for public folders
$publicFoldersSheet = $excelPackage.Workbook.Worksheets.Add("PublicFolders")

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
$publicFolders | Export-Excel -ExcelPackage $excelPackage -WorksheetName "PublicFolders" -Title "Public Folders and Child Folders" -BoldTopRow -PassThru | Out-Null

# List of distribution lists with members
Write-Host "Fetching Distribution Group lists with members..."
$distributionLists = Get-DistributionGroup | foreach {
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
$distributionLists | Export-Excel -ExcelPackage $excelPackage -WorksheetName $worksheets[3] -Title "Distribution Lists with Members" -BoldTopRow -PassThru  | Out-Null


# List of all Teams and Groups with their members
$teamsGroups = @()
Write-Host "Fetching Teams and Groups lists with members..."


# Get all Microsoft 365 groups
$teamsAndGroups = Get-AzureADGroup -All $true | 
                    Where-Object { $_.GroupTypes -contains "Unified" } | 
                    Select-Object DisplayName, Description, ObjectId

# Create an array to store teams and groups information
$teamsAndGroupsWithMembers = @()

# Iterate through each team/group
foreach ($group in $teamsAndGroups) {
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

# Check if a worksheet named "Teams_And_Groups" already exists, and remove it if it does
$existingWorksheet = $excelPackage.Workbook.Worksheets["Teams_And_Groups"]
if ($existingWorksheet) {
    $excelPackage.Workbook.Worksheets.Delete($existingWorksheet)
}

# Add a new worksheet for teams and groups
$teamsAndGroupsSheet = $excelPackage.Workbook.Worksheets.Add("Teams_And_Groups")

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

# Create an array to store security group information
$securityGroupsWithMembers = @()

# Iterate through each security group
foreach ($group in $securityGroups) {
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

# Check if a worksheet named "Security_Groups" already exists, and remove it if it does
$existingWorksheet = $excelPackage.Workbook.Worksheets["Security_Groups"]
if ($existingWorksheet) {
    $excelPackage.Workbook.Worksheets.Delete($existingWorksheet)
}

# Add a new worksheet for security groups
$securityGroupsSheet = $excelPackage.Workbook.Worksheets.Add("Security_Groups")

# Add headers
$securityGroupsSheet.Cells[1, 1].Value = "Group Display Name"
$securityGroupsSheet.Cells[1, 2].Value = "Group Member"
$securityGroupsSheet.Cells[1, 3].Value = "Group Description"

# Fill the sheet with the security groups data
$row = 2
foreach ($group in $securityGroupsWithMembers) {
    $securityGroupsSheet.Cells[$row, 1].Value = $group.GroupDisplayName
    $securityGroupsSheet.Cells[$row, 2].Value = $group.GroupMember
    $securityGroupsSheet.Cells[$row, 3].Value = $group.GroupDescription
    $row++
}


# Save the Excel package
$excelPackage.SaveAs($filePath)

Write-Host "Excel file created: $filePath"

# Prompt to open the file
$openFile = Read-Host "Do you want to open the file now? (yes/no)"

if ($openFile -eq "yes" -or $openFile -eq "y" -or $openFile -eq "ye") {
    Invoke-Item $filePath
}

# Stop the transcript to end the logging
Stop-Transcript