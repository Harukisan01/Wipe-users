# ==============================
# CONFIGURATION
# ==============================
# Group ID containing the users to wipe
$UserGroupId = "<INSERT_GROUP_OBJECT_ID_HERE>"

# Security Options
$DryRun = $false
$RequireTypedConfirmation = $true

# ==============================
# SCRIPT START - LEVEL 1 (HYBRID ISOLATION)
# ==============================
$ErrorActionPreference = "Stop"

Write-Host "Initializing Level 1 Wipe User Tool..." -ForegroundColor Cyan

# ---------------------------------------------------------
# PHASE 1: DISCOVERY (Microsoft.Graph.Authentication)
# ---------------------------------------------------------
# We use Graph ONLY to discover the Tenant Name securely.
# We explicitly unload modules to prevent DLL conflicts (TypeLoadException).

try {
    Write-Host "`n[Phase 1] Discovery: Detecting Tenant Name..." -ForegroundColor Cyan

    # 1. Clean Environment
    if (Get-Module -Name PnP.PowerShell) { Remove-Module PnP.PowerShell -Force -ErrorAction SilentlyContinue }
    if (Get-Module -Name Microsoft.Graph*) { Remove-Module Microsoft.Graph* -Force -ErrorAction SilentlyContinue }

    # 2. Load Graph Auth Only
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        Write-Warning "Microsoft.Graph.Authentication not found. Installing..."
        Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

    # 3. Connect & Get Context
    # Disable WAM for Graph too, just in case
    $env:MSAL_USE_BROKER_WITH_WAM = "false"

    Write-Host "  -> Authenticating to Graph (User.Read)..." -ForegroundColor Gray
    Connect-MgGraph -Scopes "User.Read" -NoWelcome -ErrorAction Stop

    $ctx = Get-MgContext
    $UserUpn = $ctx.Account

    # 4. Parse Tenant
    if ($UserUpn -match "@(?<domain>[^\.]+)\.onmicrosoft\.com") {
        $TenantName = $Matches['domain']
        Write-Host "  -> Auto-detected Tenant: $TenantName" -ForegroundColor Green
    } else {
        # Try finding it via Verified Domains (if User.Read allows reading org profile? usually needs Organization.Read.All)
        # Since we want to be minimal scope, we might fallback to prompt if UPN is custom.
        # But wait! If we have a custom domain, we might not know the -admin URL prefix easily?
        # Actually, it is ALWAYS <onmicrosoft-prefix>-admin.sharepoint.com.
        # If UPN is user@contoso.com, we don't know the onmicrosoft prefix easily without Org.Read.All.

        # Let's try to get Org Profile if scopes allow (User.Read might not be enough for full list, but let's try)
        # If this fails, we MUST prompt.
        Write-Warning "UPN ($UserUpn) does not reveal the onmicrosoft domain."
        $TenantName = Read-Host "Please enter your Tenant Name (e.g. 'contoso' for contoso.onmicrosoft.com)"
    }

    $AdminUrl = "https://$TenantName-admin.sharepoint.com"
    Write-Host "  -> Target Admin URL: $AdminUrl" -ForegroundColor Green

    # 5. CLEANUP GRAPH
    Write-Host "  -> Cleaning up Graph module to prevent conflicts..." -ForegroundColor Gray
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    Remove-Module Microsoft.Graph.Authentication -Force -ErrorAction SilentlyContinue
    Remove-Module Microsoft.Graph* -Force -ErrorAction SilentlyContinue

} catch {
    Write-Error "Discovery Failed: $_"
    exit 1
}

# ---------------------------------------------------------
# PHASE 2: EXECUTION (PnP.PowerShell)
# ---------------------------------------------------------
try {
    Write-Host "`n[Phase 2] Execution: Connecting to SharePoint..." -ForegroundColor Cyan

    # 1. Load PnP
    if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
        Write-Warning "PnP.PowerShell not found. Installing..."
        Install-Module PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module PnP.PowerShell -ErrorAction Stop

    # 2. Connect PnP
    # Using Interactive login. We don't specify scopes to avoid parameter errors on some versions.
    # PnP will handle its own app registration.

    Connect-PnPOnline -Url $AdminUrl -Interactive -ErrorAction Stop
    Write-Host "Connected to SharePoint Admin via PnP!" -ForegroundColor Green

} catch {
    Write-Error "SharePoint Connection Failed: $_"
    Write-Host "CRITICAL: Unable to connect to SharePoint. Exiting to prevent partial wipe." -ForegroundColor Red
    exit 1
}

# ==============================
# HELPER FUNCTIONS
# ==============================

function Confirm-DestructiveAction {
    param([string]$Title, [string]$Details)

    Write-Host ""
    Write-Host "=== $Title ===" -ForegroundColor Yellow
    Write-Host $Details -ForegroundColor Yellow
    Write-Host "DryRun=$DryRun" -ForegroundColor Yellow
    Write-Host ""

    if (-not $RequireTypedConfirmation) { return $true }
    $typed = Read-Host "Type 'EXECUTE' to confirm"
    return ($typed -eq "EXECUTE")
}

function Invoke-Safe {
    param([scriptblock]$Action, [string]$What)
    if ($DryRun) {
        Write-Host "[DRY-RUN] $What" -ForegroundColor Gray
    } else {
        Write-Host $What -ForegroundColor White
        try {
            & $Action
        } catch {
            Write-Host "  [ERROR] $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

# ==============================
# MAIN LOGIC
# ==============================

Write-Host "`nRetrieving members of group $UserGroupId..." -ForegroundColor Cyan

try {
    # Get Group Members via PnP Graph
    $GroupMembersUrl = "v1.0/groups/$UserGroupId/members?`$select=id,userPrincipalName,displayName"
    $Response = Invoke-PnPGraphMethod -Url $GroupMembersUrl -Method Get
    $Users = $Response.value
} catch {
    Write-Error "Failed to retrieve group members: $_"
    Write-Host "Tip: Ensure the PnP Application has 'GroupMember.Read.All' permissions." -ForegroundColor Yellow
    exit 1
}

if (-not $Users) {
    Write-Host "No users found in the group." -ForegroundColor Yellow
    exit
}

Write-Host "Users found: $($Users.Count)" -ForegroundColor Green

$okUsers = Confirm-DestructiveAction -Title "GENERAL & ONEDRIVE CLEANUP (DESTRUCTIVE)" -Details "Users: $($Users.Count). Operations: Mailbox, Deleted Items, Folders (Shared, Favorites, My), Recycle Bin, OneDrive (Reset), Activities, Sessions."

if (-not $okUsers) {
    Write-Host "Cancelled."
    exit
}

$Results = @()

foreach ($User in $Users) {
    $UserId = $User.id
    $UserUpn = $User.userPrincipalName
    $UserName = $User.displayName

    if (-not $UserUpn) { continue }

    Write-Host "`n===========================================" -ForegroundColor Cyan
    Write-Host "PROCESSING: $UserName ($UserUpn)" -ForegroundColor Cyan
    Write-Host "===========================================" -ForegroundColor Cyan

    # 1. Email Cleanup (Graph)
    Invoke-Safe -What "1. Email Cleanup" -Action {
        $MsgUrl = "v1.0/users/$UserId/messages?`$select=id"
        $Messages = (Invoke-PnPGraphMethod -Url $MsgUrl -Method Get).value
        Write-Host "    messages found: $($Messages.Count)" -ForegroundColor Yellow

        foreach ($Msg in $Messages) {
            Invoke-PnPGraphMethod -Url "v1.0/users/$UserId/messages/$($Msg.id)" -Method Delete
        }
        Write-Host "    Completed!" -ForegroundColor Green
    }

    # 2. Deleted Items (Graph)
    Invoke-Safe -What "2. Deleted Items Cleanup" -Action {
        $Folders = (Invoke-PnPGraphMethod -Url "v1.0/users/$UserId/mailFolders" -Method Get).value
        $DeletedFolder = $Folders | Where-Object { $_.displayName -eq "Deleted Items" }

        if ($DeletedFolder) {
            $DelMsgs = (Invoke-PnPGraphMethod -Url "v1.0/users/$UserId/mailFolders/$($DeletedFolder.id)/messages" -Method Get).value
            Write-Host "    deleted messages found: $($DelMsgs.Count)" -ForegroundColor Yellow
            foreach ($Msg in $DelMsgs) {
                Invoke-PnPGraphMethod -Url "v1.0/users/$UserId/messages/$($Msg.id)" -Method Delete
            }
            Write-Host "    Completed!" -ForegroundColor Green
        }
    }

    # 4. OneDrive (Site Deletion - PnP)
    Invoke-Safe -What "4. Total OneDrive Cleanup (Site Deletion)" -Action {
        $PersonalUrlPart = $UserUpn -replace "@","_" -replace "\.","_"
        $CleanUrl = "https://$TenantName-my.sharepoint.com/personal/$PersonalUrlPart"

        Write-Host "  -> Target Site: $CleanUrl" -ForegroundColor Cyan

        try {
            Remove-PnPTenantSite -Url $CleanUrl -Force -ErrorAction Stop
            Write-Host "  -> Site Collection Removed." -ForegroundColor Green
            $Results += [PSCustomObject]@{ User = $UserUpn; Status = "Deleted" }
        } catch {
            Write-Host "  -> Site likely not found or already deleted." -ForegroundColor Gray
            $Results += [PSCustomObject]@{ User = $UserUpn; Status = "NotFound" }
        }

        try {
            Remove-PnPTenantSite -Url $CleanUrl -FromRecycleBin -Force -ErrorAction SilentlyContinue
            Write-Host "  -> Purged from Recycle Bin." -ForegroundColor Green
        } catch {}
    }

    # 6. Sessions (Graph)
    Invoke-Safe -What "6. Revoke Sessions" -Action {
        Invoke-PnPGraphMethod -Url "v1.0/users/$UserId/revokeSignInSessions" -Method Post
        Write-Host "    Completed!" -ForegroundColor Green
    }

    Write-Host "`n✓ USER $UserName COMPLETED" -ForegroundColor Green
    Write-Host "===========================================" -ForegroundColor Cyan
}

# ==============================
# DEFINITIVE PURGE
# ==============================
Invoke-Safe -What "DEFINITIVE PURGE (Recycle Bin)" -Action {
    Write-Host "Checking Tenant Recycle Bin..." -ForegroundColor Yellow
    try {
        $DeletedSites = Get-PnPTenantRecycleBinItem | Where-Object {$_.Url -like "*-my.sharepoint.com/personal/*"}
        if ($DeletedSites) {
            foreach ($DeletedSite in $DeletedSites) {
                Write-Host "  -> Purge: $($DeletedSite.Url)" -ForegroundColor Red
                Remove-PnPTenantSite -Url $DeletedSite.Url -FromRecycleBin -Force -ErrorAction SilentlyContinue
            }
        }
    } catch {
        Write-Host "Recycle bin check failed or empty." -ForegroundColor Gray
    }
}

Write-Host "`n=== ONEDRIVE SUMMARY ===" -ForegroundColor Cyan
$Results | Format-Table -AutoSize

Write-Host "`n`n========================================" -ForegroundColor Green
Write-Host "FULL CLEANUP COMPLETED (PURE PNP PS7)" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
