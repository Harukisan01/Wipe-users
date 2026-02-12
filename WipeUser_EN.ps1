 ==============================================================================

# --- 1. CONFIGURATION ---
$UserGroupId = "YOUR_GROUP_ID_HERE" 
$DryRun = $false 
$RequireTypedConfirmation = $true 

# --- 2. INITIALIZATION ---
$ErrorActionPreference = "Stop" 
$Scopes = @( 
    "GroupMember.Read.All", "Mail.ReadWrite", "Files.ReadWrite.All", 
    "User.ReadWrite.All", "Sites.ReadWrite.All", "Sites.FullControl.All", 
    "Organization.Read.All", "Contacts.ReadWrite", "Calendars.ReadWrite", "Tasks.ReadWrite"
) 

# --- 3. AUTHENTICATION ---
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan 
Connect-MgGraph -Scopes $Scopes -NoWelcome 
$ctx = Get-MgContext 

# SharePoint Admin Connection
try { 
    $Org = Get-MgOrganization 
    $TenantName = ($Org.VerifiedDomains | Where-Object { $_.Name -like "*.onmicrosoft.com" } | Select-Object -First 1).Name -replace "\.onmicrosoft\.com", "" 
    $AdminUrl = "https://$TenantName-admin.sharepoint.com" 
    Connect-SPOService -Url $AdminUrl 
} catch { 
    Write-Warning "Could not connect to SharePoint Admin. Site deletion might fail." 
} 

# ============================== 
# FUNCTIONS 
# ============================== 

function Purge-MailboxData {
    param($UserId)
    Write-Host "  -> Purging Mailbox (Email, Calendar, Contacts, Tasks)..." -ForegroundColor Yellow
    
    # 1. Emails
    $Msgs = Get-MgUserMessage -UserId $UserId -All -Property Id
    foreach ($M in $Msgs) { Remove-MgUserMessage -UserId $UserId -MessageId $M.Id -Confirm:$false }
    
    # 2. Contacts
    $Contacts = Get-MgUserContact -UserId $UserId -All -Property Id
    foreach ($C in $Contacts) { Remove-MgUserContact -UserId $UserId -ContactId $C.Id -Confirm:$false }
    
    # 3. Calendar Events
    $Events = Get-MgUserEvent -UserId $UserId -All -Property Id
    foreach ($E in $Events) { Remove-MgUserEvent -UserId $UserId -EventId $E.Id -Confirm:$false }
}

function Clear-SharedWithMeView {
    param($UserId)
    Write-Host "  -> Revoking access to 'Shared with me' items..." -ForegroundColor Yellow
    $SharedItems = Get-MgUserDriveSharedWithMe -UserId $UserId -ErrorAction SilentlyContinue
    foreach ($Item in $SharedItems) {
        try {
            $OwnerDriveId = $Item.RemoteItem.DriveId
            $ItemId = $Item.RemoteItem.Id
            $Perms = Get-MgDriveItemPermission -DriveId $OwnerDriveId -DriveItemId $ItemId -ErrorAction SilentlyContinue
            $TargetPerm = $Perms | Where-Object { $_.GrantedTo.User.Id -eq $UserId -or $_.GrantedToV2.User.Id -eq $UserId }
            foreach ($P in $TargetPerm) { Remove-MgDriveItemPermission -DriveId $OwnerDriveId -DriveItemId $ItemId -PermissionId $P.Id -Confirm:$false -ErrorAction SilentlyContinue }
        } catch { continue }
    }
}

function Invoke-Safe {
    param([scriptblock]$Action, [string]$What)
    if ($DryRun) { Write-Host "[DRY-RUN] Skipping: $What" -ForegroundColor Gray } 
    else { Write-Host ">>> $What" -ForegroundColor White; & $Action }
}

# ============================== 
# MAIN EXECUTION 
# ============================== 

$Users = Get-MgGroupMember -GroupId $UserGroupId -All | Where-Object { $_.AdditionalProperties.'@odata.type' -eq "#microsoft.graph.user" } 

if ($RequireTypedConfirmation) {
    $Confirm = Read-Host "Type 'DELETE EVERYTHING' to confirm for $($Users.Count) users"
    if ($Confirm -ne "DELETE EVERYTHING") { Write-Host "Aborted."; exit }
}

foreach ($UserRef in $Users) { 
    $UserId = $UserRef.Id 
    $User = Get-MgUser -UserId $UserId -Property UserPrincipalName,DisplayName
    $UserUpn = $User.UserPrincipalName
    Write-Host "`n>>> Wiping Everything for: $($User.DisplayName)" -ForegroundColor Cyan 

    # 1. TOTAL MAILBOX PURGE
    Invoke-Safe -What "Full Mailbox Wipe" -Action { Purge-MailboxData -UserId $UserId }

    # 2. ONEDRIVE & SHARED LIST
    Invoke-Safe -What "Clearing Shared List & Permissions" -Action { Clear-SharedWithMeView -UserId $UserId }

    # 3. NUCLEAR SITE DELETE & RESET
    Invoke-Safe -What "Deleting OneDrive Site Collection" -Action { 
        $PersonalUrl = "https://$TenantName-my.sharepoint.com/personal/$($UserUpn -replace '[\.@]', '_')"
        Remove-SPOSite -Identity $PersonalUrl -NoWait -Confirm:$false -ErrorAction SilentlyContinue
        Remove-SPODeletedSite -Identity $PersonalUrl -NoWait -Confirm:$false -ErrorAction SilentlyContinue
        Request-SPOPersonalSite -UserEmails $UserUpn -NoWait
    }

    # 4. SESSIONS & HISTORY
    Invoke-Safe -What "Revoking Sessions & History" -Action { 
        Revoke-MgUserSignInSession -UserId $UserId 
        $Acts = Get-MgUserActivity -UserId $UserId -All
        foreach ($A in $Acts) { Remove-MgUserActivity -UserId $UserId -ActivityId $A.Id -Confirm:$false }
    }
}

Write-Host "`nPURGE COMPLETE: All data has been destroyed." -ForegroundColor Green
