# Import the Exchange Online module (install if needed)
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber
}

# Import the module
Import-Module ExchangeOnlineManagement

# Log in to Exchange Online
Write-Host "Logging in to Exchange Online..." -ForegroundColor Yellow
$session = Connect-ExchangeOnline -ShowProgress $true

# Prompt for the user's email address
$userEmail = Read-Host "Enter the email address of the user"

function List-Rules {
    # Retrieve and display all inbox rules, including hidden ones, for the specified user
    Write-Host "`nFetching inbox rules (including hidden rules) for ${userEmail}..." -ForegroundColor Cyan
    $rules = Get-InboxRule -Mailbox $userEmail -IncludeHidden

    if ($rules.Count -eq 0) {
        Write-Host "No inbox rules found for ${userEmail}." -ForegroundColor Red
        return $null
    }

    Write-Host "`nDetailed Inbox Rules for ${userEmail} (including hidden rules):" -ForegroundColor Green
    $rules | Select-Object Name, Description, Enabled, RedirectTo, MoveToFolder, ForwardTo | Format-List
    return $rules
}

# Initial listing of rules
$rules = List-Rules
if (-not $rules) {
    Disconnect-ExchangeOnline -Confirm:$false
    exit
}

# Allow user to manage rules
while ($true) {
    Write-Host "`nWhat would you like to do?"
    Write-Host "1. Enable a rule"
    Write-Host "2. Disable a rule"
    Write-Host "3. Delete a rule"
    Write-Host "4. List rules again"
    Write-Host "5. Exit"

    $choice = Read-Host "Enter your choice (1/2/3/4/5)"

    switch ($choice) {
        1 {
            $ruleID = Read-Host "Enter the ID of the rule to enable"
            Set-InboxRule -Mailbox $userEmail -Identity $ruleID -Enabled $true
            Write-Host "Rule [$ruleID] has been enabled." -ForegroundColor Green
        }
        2 {
            $ruleID = Read-Host "Enter the ID of the rule to disable"
            Set-InboxRule -Mailbox $userEmail -Identity $ruleID -Enabled $false
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

            Remove-InboxRule -Mailbox $userEmail -Identity $ruleID -Confirm:$false
            Write-Host "Rule [$ruleID] has been PERMANENTLY deleted." -ForegroundColor Green
        }
        4 {
            # Re-list the rules
            $rules = List-Rules
        }
        5 {
            Write-Host "Exiting..." -ForegroundColor Yellow
            break
        }
        Default {
            Write-Host "Invalid choice. Please try again." -ForegroundColor Red
        }
    }
}

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false
