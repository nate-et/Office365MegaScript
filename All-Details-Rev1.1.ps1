# Install ImportExcel module if not already installed
if (-not (Get-Module -Name ImportExcel -ListAvailable)) {
    Install-Module -Name ImportExcel -Scope CurrentUser -Force -Verbose
}

# Import the module
Import-Module ImportExcel

# Prompt for Exchange Online credentials
$credential = Get-Credential -Message "Enter your Exchange Online admin credentials"

# Connect to Required Modules
Connect-ExchangeOnline -Credential $credential
Connect-AzureAD -Credential $credential
Connect-MsolService -Credential $credential


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
$licensedAccounts = Get-MsolUser -All | where {$_.isLicensed -eq "True"} | foreach {
    [PSCustomObject]@{
        "UPN" = $_.UserPrincipalName
        "License" = ($_.Licenses | ForEach-Object { $_.AccountSkuId }) -join ', '  # Convert array to comma-separated string
    }
}

# Export licensed accounts to Excel
$licensedAccounts | Export-Excel -ExcelPackage $excelPackage -WorksheetName $worksheets[1] -Title "Licensed Accounts" -BoldTopRow -PassThru  | Out-Null

# List of shared mailboxes with delegated mailbox access
$sharedMailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox | foreach {
    $mailbox = $_.DisplayName
    $delegates = Get-MailboxPermission -Identity $_.Identity -User $_.UserPrincipalName -ErrorAction SilentlyContinue
    if ($delegates) {
        $delegates | foreach {
            [PSCustomObject]@{
                "SharedMailbox" = $mailbox
                "DelegateName" = $_.User.DisplayName
                "DelegateUPN" = $_.User.UserPrincipalName
            }
        }
    }
}

# Export shared mailboxes with delegated mailbox access to Excel
$sharedMailboxes | Export-Excel -ExcelPackage $excelPackage -WorksheetName $worksheets[2] -Title "Shared Mailboxes with Delegated Access" -BoldTopRow -PassThru  | Out-Null

# List of distribution lists with members
$distributionLists = Get-DistributionGroup | foreach {
    $groupName = $_.DisplayName
    $members = Get-DistributionGroupMember -Identity $_.Identity | foreach {
        [PSCustomObject]@{
            "DistributionList" = $groupName
            "MemberName" = $_.DisplayName
            "MemberEmail" = $_.PrimarySmtpAddress
        }
    }
}

# Export distribution lists with members to Excel
$distributionLists | Export-Excel -ExcelPackage $excelPackage -WorksheetName $worksheets[3] -Title "Distribution Lists with Members" -BoldTopRow -PassThru  | Out-Null

# List of all Teams and Groups with their members
$teamsGroups = Get-AzureADGroup | foreach {
    $groupName = $_.DisplayName
    $groupMembers = Get-AzureADGroupMember -ObjectId $_.ObjectId | foreach {
        [PSCustomObject]@{
            "Group/Team" = $groupName
            "MemberName" = $_.DisplayName
            "MemberUPN" = $_.UserPrincipalName
        }
    }
}

# Export Teams and Groups with members to Excel
$teamsGroups | Export-Excel -ExcelPackage $excelPackage -WorksheetName $worksheets[4] -Title "Teams and Groups with Members" -BoldTopRow -PassThru | Out-Null

# Save the Excel package
$excelPackage.SaveAs($filePath)

Write-Host "Excel file created: $filePath"