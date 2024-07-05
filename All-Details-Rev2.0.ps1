# Notes: To come in future release:
# 1. Security Groups
# 2. Better Definitions for "Type" of teams group
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
    "TeamsGroups_Members"
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
Write-Host "Fetching distribution lists with members..."
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

    foreach ($member in $members) {
        $groupType = if ($group.GroupTypes -contains "Unified") { "Teams-Enabled Group" } else { "Regular Group" }

        $teamsGroups += [PSCustomObject]@{
            "Group/Team" = $group.DisplayName
            "GroupType" = $groupType
            "MemberName" = $member.DisplayName
            "MemberUPN" = $member.PrimarySmtpAddress
        }
    }
}


# Export Teams and Groups with members to Excel
$teamsGroups | Export-Excel -ExcelPackage $excelPackage -WorksheetName $worksheets[4] -Title "Teams and Groups with Members" -BoldTopRow -PassThru | Out-Null

# Save the Excel package
$excelPackage.SaveAs($filePath)

Write-Host "Excel file created: $filePath"