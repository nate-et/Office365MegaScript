# Import the Exchange Online module (install if needed)
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber
}

# Import the module
Import-Module ExchangeOnlineManagement

# Log in to Exchange Online
Write-Host "Logging in to Exchange Online..." -ForegroundColor Yellow
$session = Connect-ExchangeOnline -ShowProgress $true

# Function to prompt for a user and fetch their rules
function Select-User {
    param (
        [string]$CurrentUser = ""
    )
    if ($CurrentUser) {
        Write-Host "`nCurrent selected user: $CurrentUser" -ForegroundColor Cyan
    }
    $newUserEmail = Read-Host "Enter the email address of the user you want to manage"
    return $newUserEmail
}

# Function to fetch and display inbox rules, with error handling for any corrupt rules
function List-Rules {
    param (
        [string]$Mailbox
    )
    Write-Host "`nFetching inbox rules (including hidden rules) for ${Mailbox}..." -ForegroundColor Cyan
    $rules = Get-InboxRule -Mailbox $Mailbox -IncludeHidden

    if ($rules.Count -eq 0) {
        Write-Host "No inbox rules found for ${Mailbox}." -ForegroundColor Red
        return $null
    }

    Write-Host "`nDetailed Inbox Rules for ${Mailbox} (including hidden rules):" -ForegroundColor Green
    foreach ($rule in $rules) {
        try {
            # Display each rule and handle errors for specific problematic rules
            $rule | Select-Object Name, Description, Enabled, RedirectTo, MoveToFolder, ForwardTo | Format-List
        }
        catch {
            Write-Host "Error fetching rule '$($rule.Name)': $_" -ForegroundColor Red
        }
    }

    return $rules
}

# Initialize the selected user
$currentUser = Select-User
$rules = List-Rules -Mailbox $currentUser
if (-not $rules) {
    Disconnect-ExchangeOnline -Confirm:$false
    exit
}

# Persistent loop for managing rules
while ($true) {
    Write-Host "`nCurrent selected user: $currentUser" -ForegroundColor Cyan
    Write-Host "What would you like to do?"
    Write-Host "1. Enable a rule"
    Write-Host "2. Disable a rule"
    Write-Host "3. Delete a rule"
    Write-Host "4. List rules again"
    Write-Host "5. Change user"
    Write-Host "6. Exit"

    $choice = Read-Host "Enter your choice (1/2/3/4/5/6)"

    switch ($choice) {
        1 {
            $ruleID = Read-Host "Enter the ID of the rule to enable"
            Set-InboxRule -Mailbox $currentUser -Identity $ruleID -Enabled $true
            Write-Host "Rule [$ruleID] has been enabled." -ForegroundColor Green
        }
        2 {
            $ruleID = Read-Host "Enter the ID of the rule to disable"
            Set-InboxRule -Mailbox $currentUser -Identity $ruleID -Enabled $false
            Write-Host "Rule [$ruleID] has been disabled." -ForegroundColor Green
        }
        3 {
            Write-Host "`nWARNING: Are you sure you want to delete a rule? This action is IRREVERSIBLE and should only be used if you cannot disable the rule because it is corrupted." -ForegroundColor Yellow
            $confirm1 = Read-Host "Type 'Yes' to proceed or 'No' to return to the main menu"
            if ($confirm1 -ne "Yes") {
                Write-Host "Returning to main menu..." -ForegroundColor Cyan
                continue
            }

            $ruleID = Read-Host "Enter the ID of the rule to delete"
            Write-Host "`nAre you CERTAIN you want to PERMANENTLY delete rule [$ruleID]? This action is IRREVERSIBLE!" -ForegroundColor Red
            $confirm2 = Read-Host "Type 'Delete' to confirm or 'Cancel' to return to the main menu"
            if ($confirm2 -ne "Delete") {
                Write-Host "Rule deletion cancelled. Returning to main menu..." -ForegroundColor Cyan
                continue
            }

            Remove-InboxRule -Mailbox $currentUser -Identity $ruleID -Confirm:$false
            Write-Host "Rule [$ruleID] has been PERMANENTLY deleted." -ForegroundColor Green
        }
        4 {
            # Re-list the rules
            $rules = List-Rules -Mailbox $currentUser
        }
        5 {
            # Change the selected user
            $currentUser = Select-User -CurrentUser $currentUser
            $rules = List-Rules -Mailbox $currentUser
            if (-not $rules) {
                Disconnect-ExchangeOnline -Confirm:$false
                exit
            }
        }
        6 {
            Write-Host "Exiting..." -ForegroundColor Yellow
            Disconnect-ExchangeOnline -Confirm:$false
            exit
        }
        Default {
            Write-Host "Invalid choice. Please try again." -ForegroundColor Red
        }
    }
}
