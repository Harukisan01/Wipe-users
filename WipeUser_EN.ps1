$UserGroupId = ""

$DryRun = $false
$RequireTypedConfirmation = $true

$Scopes = @(
    "GroupMember.Read.All",
    "Mail.ReadWrite",
    "Files.ReadWrite.All",
    "User.ReadWrite.All"
)

Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
Write-Host "Login to Microsoft Graph (interactive)..."
Connect-MgGraph -Scopes $Scopes -NoWelcome

$ctx = Get-MgContext
Write-Host "Connected as: $($ctx.Account) | Tenant: $($ctx.TenantId)"

function Test-GuidOrThrow {
    param([string]$Value, [string]$Name)
    $g = [guid]::Empty
    if (-not [guid]::TryParse($Value, [ref]$g)) {
        throw "$Name is not a valid GUID: '$Value'"
    }
}

function Confirm-DestructiveAction {
    param([string]$Title, [string]$Details)
    Write-Host ""
    Write-Host "=== $Title ===" -ForegroundColor Yellow
    Write-Host $Details -ForegroundColor Yellow
    Write-Host "DryRun=$DryRun" -ForegroundColor Yellow
    Write-Host ""
    if (-not $RequireTypedConfirmation) { return $true }
    $typed = Read-Host "Type EXACTLY 'EXECUTE' to continue (anything else = cancel)"
    return ($typed -eq "EXECUTE")
}

function Invoke-Safe {
    param([scriptblock]$Action, [string]$What)
    if ($DryRun) {
        Write-Host "[DRY-RUN] $What" -ForegroundColor Gray
    } else {
        Write-Host $What -ForegroundColor White
        & $Action
    }
}

Test-GuidOrThrow -Value $UserGroupId -Name "UserGroupId"

Write-Host "`nRetrieving group members..."
$Users = Get-MgGroupMember -GroupId $UserGroupId -All |
    Where-Object { $_.AdditionalProperties.'@odata.type' -eq "#microsoft.graph.user" }

Write-Host "Users found: $($Users.Count)"

if ($Users.Count -gt 0) {
    $okUsers = Confirm-DestructiveAction `
        -Title "USER OPERATIONS (DESTRUCTIVE)" `
        -Details "For $($Users.Count) users:`n- Delete emails (mailbox)`n- Delete Deleted Items`n- Delete OneDrive (Root + Recycle Bin)`n- Delete Activities (Timeline)`n- Revoke sessions"

    if (-not $okUsers) {
        Write-Host "User operations cancelled."
    } else {
        foreach ($UserRef in $Users) {
            $UserId = $UserRef.Id
            Write-Host "`n--- User: $UserId ---" -ForegroundColor Cyan

            Invoke-Safe -What "Deleting mailbox messages for $UserId" -Action {
                $Messages = Get-MgUserMessage -UserId $UserId -All -Property Id
                foreach ($Msg in $Messages) {
                    Remove-MgUserMessage -UserId $UserId -MessageId $Msg.Id -Confirm:$false
                }
            }

            Invoke-Safe -What "Deleting 'Deleted Items' for $UserId" -Action {
                $Deleted = Get-MgUserMailFolder -UserId $UserId -All |
                    Where-Object { $_.DisplayName -eq "Deleted Items" } |
                    Select-Object -First 1
                if ($Deleted) {
                    $DeletedMessages = Get-MgUserMailFolderMessage -UserId $UserId -MailFolderId $Deleted.Id -All -Property Id
                    foreach ($Msg in $DeletedMessages) {
                        Remove-MgUserMailFolderMessage -UserId $UserId -MailFolderId $Deleted.Id -MessageId $Msg.Id -Confirm:$false
                    }
                }
            }

            Invoke-Safe -What "Cleaning OneDrive (Root) and Recycle Bin for $UserId" -Action {
                try {
                    $drive = Get-MgUserDrive -UserId $UserId -Property Id -ErrorAction Stop
                } catch {
                    if ($_.Exception.Message -match "Access denied") {
                        Write-Host "  [!] ACCESS DENIED: You do not have access to this user's OneDrive." -ForegroundColor Red
                        Write-Host "  [?] SOLUTION: Add your account as 'Site Collection Admin' for this user." -ForegroundColor Yellow
                        return
                    } else {
                        Write-Host "  [!] Error retrieving OneDrive: $_" -ForegroundColor Red
                        return
                    }
                }

                if ($drive) {
                    try {
                        $items = Get-MgDriveRootChild -DriveId $drive.Id -All -Property Id, Name -ErrorAction Stop
                        foreach ($item in $items) {
                            try {
                                Remove-MgDriveItem -DriveId $drive.Id -DriveItemId $item.Id -Confirm:$false -ErrorAction Stop
                            } catch { Write-Host "    [!] Error deleting file '$($item.Name)': $_" -ForegroundColor Red }
                        }
                    } catch { Write-Host "  [!] Error reading Root: $_" -ForegroundColor Red }

                    try {
                        $binItems = Get-MgDriveRecycleBin -DriveId $drive.Id -All -ErrorAction SilentlyContinue
                        foreach ($binItem in $binItems) {
                            Write-Host "  -> Permanently deleting: $($binItem.Name)" -ForegroundColor DarkGray
                            Invoke-MgGraphRequest -Method DELETE -Uri "drives/$($drive.Id)/recycleBin/$($binItem.Id)"
                        }
                    } catch { Write-Host "  [!] Error accessing OneDrive Recycle Bin: $_" -ForegroundColor Red }
                }
            }

            Invoke-Safe -What "Deleting User Activities (Timeline/History) for $UserId" -Action {
                try {
                    $activities = Get-MgUserActivity -UserId $UserId -All -ErrorAction SilentlyContinue
                    foreach ($act in $activities) {
                        Remove-MgUserActivity -UserId $UserId -ActivityId $act.Id -Confirm:$false
                    }
                } catch { Write-Host "  [!] Unable to clean activities (permissions or feature missing): $_" -ForegroundColor DarkGray }
            }

            Invoke-Safe -What "Revoking sessions for $UserId" -Action {
                Revoke-MgUserSignInSession -UserId $UserId | Out-Null
            }
        }
    }
} else {
    Write-Host "No users in group."
}

Write-Host "`nOPERATIONS COMPLETE."
