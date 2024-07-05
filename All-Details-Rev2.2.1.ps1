# Notes: To come in future release:
# 
# 1. Better Definitions for "Type" of teams group
#
#
# ------------------------------------------------------------------------------------------------------------
#
#
# Change List
#
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
$credential = Get-Credential -Message "Enter your Exchange Online admin credentials"


Write-Host "Connecting To Exchange Online, AzureAD & MsolService..."
# Connect to Required Modules
Connect-ExchangeOnline -Credential $credential
Connect-AzureAD -Credential $credential
Connect-MsolService -Credential $credential

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


# Get all Office 365 groups
$groups = Get-UnifiedGroup -ResultSize Unlimited

foreach ($group in $groups) {
    # Check if the group is a distribution group or a security group
    if ($group.RecipientTypeDetails -eq "MailUniversalDistributionGroup" -or $group.RecipientTypeDetails -eq "MailUniversalSecurityGroup") {
        $members = Get-DistributionGroupMember -Identity $group.Identity
    } elseif ($group.RecipientTypeDetails -eq "MailNonUniversalGroup") {
        $members = Get-UnifiedGroupLinks -Identity $group.Identity -LinkType Members
    }

    # Determine the privacy setting based on AccessType
    $privacy = if ($group.AccessType -eq "Public") { "Public" } elseif ($group.AccessType -eq "Private") { "Private" } else { "Unknown" }

    foreach ($member in $members) {
        $teamsGroups += [PSCustomObject]@{
            "Group/Team" = $group.DisplayName
            "Privacy" = $privacy
            "MemberName" = $member.DisplayName
            "MemberUPN" = $member.PrimarySmtpAddress
        }
    }
}


# Export Teams and Groups with members to Excel
$teamsGroups | Export-Excel -ExcelPackage $excelPackage -WorksheetName $worksheets[4] -Title "Teams and Groups with Members" -BoldTopRow -PassThru | Out-Null


Write-Host "Fetching Security Groups list with members..."

# Get all security groups with DisplayName and Description
$securityGroups = Get-AzureADGroup -All $true | 
                    Where-Object { $_.SecurityEnabled -eq $true } | 
                    Select-Object DisplayName, Description

# Create an array to store security group members
$securityGroupsWithMembers = @()

# Iterate through each security group
foreach ($group in $securityGroups) {
    # Fetch members of the current security group
    $groupMembers = Get-AzureADGroupMember -ObjectId $group.ObjectId | 
                    Select-Object DisplayName, UserPrincipalName
    
    # Convert group members to a comma-separated string for better readability
    $members = $groupMembers.DisplayName -join ", "

    # Create an object to store group information along with its members
    $groupObject = [PSCustomObject]@{
        "GroupDisplayName" = $group.DisplayName
        "GroupDescription" = $group.Description
        "GroupMembers" = $members
    }

    # Add the group object to the array
    $securityGroupsWithMembers += $groupObject
}

# Export Security Groups to Excel without specific properties
$securityGroups | Export-Excel -ExcelPackage $excelPackage -WorksheetName $worksheets[5] -Title "Security Groups" -BoldTopRow -PassThru | Out-Null


# Save the Excel package
$excelPackage.SaveAs($filePath)

Write-Host "Excel file created: $filePath"

# Prompt to open the file
$openFile = Read-Host "Do you want to open the file now? (yes/no)"

if ($openFile -eq "yes" -or $openFile -eq "y") {
    Invoke-Item $filePath
}