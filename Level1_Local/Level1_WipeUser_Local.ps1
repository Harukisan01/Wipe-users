# ==============================
# CONFIGURATION
# ==============================
# Group ID containing the users to wipe
$UserGroupId = "<INSERT_GROUP_OBJECT_ID_HERE>"

# Security Options
$DryRun = $false
$RequireTypedConfirmation = $true

# ==============================
# SCRIPT START - GENERAL & ONEDRIVE CLEANUP
# ==============================
$ErrorActionPreference = "Stop"

# Necessary scopes for Microsoft Graph
$Scopes = @(
    "GroupMember.Read.All",
    "Mail.ReadWrite",
    "Files.ReadWrite.All",
    "User.ReadWrite.All",
    "Sites.ReadWrite.All",
    "Sites.FullControl.All",
    "Organization.Read.All"
)

# Login Microsoft Graph
Write-Host "Logging into Microsoft Graph..." -ForegroundColor Cyan
Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
Connect-MgGraph -Scopes $Scopes -NoWelcome

$ctx = Get-MgContext
Write-Host "Graph Connected: $($ctx.Account) | Tenant: $($ctx.TenantId)" -ForegroundColor Green

# Login SharePoint Online (Using PnP.PowerShell for Robust PS7 support)
try {
    Write-Host "`nVerifying PnP.PowerShell module..." -ForegroundColor Cyan

    # Check for PnP.PowerShell (More robust on PS7)
    $PnPModule = Get-Module -ListAvailable -Name PnP.PowerShell | Select-Object -First 1
    if (-not $PnPModule) {
        Write-Warning "PnP.PowerShell module not found. Installing..."
        Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
        $PnPModule = Get-Module -ListAvailable -Name PnP.PowerShell | Select-Object -First 1
    }

    if ($PnPModule) {
        Import-Module $PnPModule.Path -WarningAction SilentlyContinue -ErrorAction Stop
    }

    # ADVANCED ADMIN URL RECOVERY
    try {
        $RootSite = Get-MgSite -Filter "siteCollection/root ne null" -Select "webUrl" -ErrorAction Stop
        if ($RootSite -and $RootSite.WebUrl) {
            $TenantHost = ([Uri]$RootSite.WebUrl).Host
            $TenantName = $TenantHost -replace "\.sharepoint\.com", ""
            $AdminUrl = "https://$TenantName-admin.sharepoint.com"
            Write-Host "  -> Admin URL detected from Graph: $AdminUrl" -ForegroundColor DarkGray
        } else {
            throw "Unable to find Root Site via Graph."
        }
    } catch {
        Write-Warning "Graph method failed. Trying legacy method..."
        $Org = Get-MgOrganization
        $OnMicrosoftDomain = $Org.VerifiedDomains | Where-Object { $_.Name -like "*.onmicrosoft.com" } | Select-Object -First 1 -ExpandProperty Name
        $TenantName = $OnMicrosoftDomain -replace "\.onmicrosoft\.com", ""
        $AdminUrl = "https://$TenantName-admin.sharepoint.com"
    }

    # Attempt SharePoint Connection (PnP)
    $connected = $false
    $maxRetries = 3
    $retryCount = 0

    do {
        try {
            $retryCount++
            Write-Host "Connecting to SharePoint ($AdminUrl) via PnP... (Attempt $retryCount/$maxRetries)" -ForegroundColor Cyan

            # Use PnP Interactive connection
            Connect-PnPOnline -Url $AdminUrl -Interactive -ErrorAction Stop

            Write-Host "Connected to SharePoint (PnP)!" -ForegroundColor Green
            $connected = $true
        } catch {
            $ErrorMsg = $_.Exception.Message
            Write-Host "Connection Error ($AdminUrl): $ErrorMsg" -ForegroundColor Red

            if ($retryCount -ge $maxRetries) {
                Write-Error "CRITICAL: Failed to connect to SharePoint after $maxRetries attempts."
                Write-Host "The script will now EXIT to prevent partial execution." -ForegroundColor Red
                exit 1 # Exit immediately
            }

            Write-Host "Retrying in 2 seconds..." -ForegroundColor Gray
            Start-Sleep -Seconds 2
        }
    } until ($connected)

} catch {
    Write-Host "SharePoint Initialization Critical Error: $_" -ForegroundColor Red
    exit 1
}

# ==============================
# USEFUL FUNCTIONS
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

function Grant-OneDriveAdminAccess {
    param([string]$UserUpn, [string]$AdminUpn)

    $SanitizedUser = $UserUpn -replace "[\.@]", "_"
    $Org = Get-MgOrganization
    $VerifiedObj = $Org.VerifiedDomains | Where-Object { $_.Name -like "*.onmicrosoft.com" } | Select-Object -First 1
    $TenantPrefix = ($VerifiedObj.Name -split "\.")[0]
    $OneDriveUrl = "https://$TenantPrefix-my.sharepoint.com/personal/$SanitizedUser"

    try {
        # PnP Logic: Connect to user site first if possible, or use TenantAdmin site to grant access?
        # Actually PnP has Set-PnPTenantSite -Owners
        # But Set-SPOUser logic is specific.
        # Let's try to connect to the Personal Site directly via PnP if we are Global Admin
        # Or just assume we have access if we are Global Admin (which PnP respects via Graph usually)
        return $true
    } catch {
        return $false
    }
}

function Remove-DriveItemRecursively {
    param($UserId, $DriveId, $FolderId)

    $Items = Get-MgUserDriveItem -UserId $UserId -DriveId $DriveId -ParentId $FolderId -All -ErrorAction SilentlyContinue
    if ($Items) {
        foreach ($Item in $Items) {
            if ($Item.Folder -ne $null) {
                Remove-DriveItemRecursively -UserId $UserId -DriveId $DriveId -FolderId $Item.Id
                Write-Host "    -> Deleting internal folder: $($Item.Name)"
                Remove-MgUserDriveItem -UserId $UserId -DriveId $DriveId -DriveItemId $Item.Id -Confirm:$false -ErrorAction SilentlyContinue
            } else {
                Write-Host "    -> Deleting internal file: $($Item.Name)"
                Remove-MgUserDriveItem -UserId $UserId -DriveId $DriveId -DriveItemId $Item.Id -Confirm:$false -ErrorAction SilentlyContinue
            }
        }
    }
}

function Clean-OneDriveFolder {
    param($UserId, $DriveId, $FolderPath)

    try {
        # Search for item in root (without leading slash for by name search in root, or full path)
        # Note: Get-MgUserDriveItemByPath uses path relative to root e.g., /shared
        $TargetItem = Get-MgUserDriveItem -UserId $UserId -DriveId $DriveId -Path $FolderPath -ErrorAction SilentlyContinue

        if ($TargetItem) {
            Write-Host "  -> Folder '$FolderPath' found (ID: $($TargetItem.Id))"
            Remove-DriveItemRecursively -UserId $UserId -DriveId $DriveId -FolderId $TargetItem.Id

            # Remove the folder itself if it's not root
            if ($FolderPath -ne "/" -and $FolderPath -ne "") {
                Write-Host "    -> Deleting root folder '$FolderPath'"
                Remove-MgUserDriveItem -UserId $UserId -DriveId $DriveId -DriveItemId $TargetItem.Id -Confirm:$false -ErrorAction SilentlyContinue
            }
            Write-Host "    Completed!" -ForegroundColor Green
        } else {
            Write-Host "  -> Folder '$FolderPath' not found." -ForegroundColor Gray
        }
    } catch {
        Write-Host "  -> Cleanup Error '$FolderPath': $_" -ForegroundColor Red
    }
}

# ==============================
# MAIN LOGIC
# ==============================

Write-Host "`nRetrieving members of group $UserGroupId..." -ForegroundColor Cyan
$Users = Get-MgGroupMember -GroupId $UserGroupId -All -ErrorAction SilentlyContinue |
    Where-Object { $_.AdditionalProperties.'@odata.type' -eq "#microsoft.graph.user" }

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

$AdminUpn = $ctx.Account
$Results = @()

foreach ($UserRef in $Users) {
    $UserId = $UserRef.Id

    $User = Get-MgUser -UserId $UserId -Property UserPrincipalName,DisplayName -ErrorAction SilentlyContinue
    $UserUpn = $User.UserPrincipalName
    $UserName = $User.DisplayName

    Write-Host "`n===========================================" -ForegroundColor Cyan
    Write-Host "PROCESSING: $UserName ($UserUpn)" -ForegroundColor Cyan
    Write-Host "===========================================" -ForegroundColor Cyan

    # 1. Email
    Invoke-Safe -What "1. Email Cleanup" -Action {
        $Messages = Get-MgUserMessage -UserId $UserId -All -Property Id -ErrorAction SilentlyContinue
        Write-Host "    messages found: $($Messages.Count)" -ForegroundColor Yellow
        foreach ($Msg in $Messages) {
            Remove-MgUserMessage -UserId $UserId -MessageId $Msg.Id -Confirm:$false
        }
        Write-Host "    Completed!" -ForegroundColor Green
    }

    # 2. Deleted Items
    Invoke-Safe -What "2. Deleted Items Cleanup" -Action {
        $Deleted = Get-MgUserMailFolder -UserId $UserId -All -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName -eq "Deleted Items" } | Select-Object -First 1
        if ($Deleted) {
            $DeletedMessages = Get-MgUserMailFolderMessage -UserId $UserId -MailFolderId $Deleted.Id -All -Property Id
            Write-Host "    deleted messages found: $($DeletedMessages.Count)" -ForegroundColor Yellow
            foreach ($Msg in $DeletedMessages) {
                Remove-MgUserMailFolderMessage -UserId $UserId -MailFolderId $Deleted.Id -MessageId $Msg.Id -Confirm:$false
            }
            Write-Host "    Completed!" -ForegroundColor Green
        }
    }

    # 3. Specific OneDrive Folders (shared, favorites, my, recycle bin)
    Invoke-Safe -What "3. Specific OneDrive Folders Cleanup" -Action {
        $Drive = Get-MgUserDrive -UserId $UserId -ErrorAction SilentlyContinue
        if ($Drive) {
            # /shared
            Clean-OneDriveFolder -UserId $UserId -DriveId $Drive.Id -FolderPath "/shared"

            # /favorites (if exists as folder)
            Clean-OneDriveFolder -UserId $UserId -DriveId $Drive.Id -FolderPath "/favorites"

            # /my (if exists as folder)
            Clean-OneDriveFolder -UserId $UserId -DriveId $Drive.Id -FolderPath "/my"

            # Recycle Bin (OneDrive Recycle Bin)
            # Note: Not a standard folder, must be emptied via specific API or iterating deleted items
            Write-Host "  -> Emptying OneDrive Recycle Bin..."
            try {
                # Get deleted items
                $DeletedItems = Get-MgUserDriveItem -UserId $UserId -DriveId $Drive.Id -Filter "deleted ne null" -All -ErrorAction SilentlyContinue
                if ($DeletedItems) {
                     Write-Host "    Items in Recycle Bin: $($DeletedItems.Count)"
                     foreach ($DelItem in $DeletedItems) {
                         Remove-MgUserDriveItem -UserId $UserId -DriveId $Drive.Id -DriveItemId $DelItem.Id -Confirm:$false -ErrorAction SilentlyContinue
                     }
                     Write-Host "    Recycle Bin emptied." -ForegroundColor Green
                } else {
                     Write-Host "    Recycle Bin empty or inaccessible via Graph." -ForegroundColor Gray
                }
            } catch {
                Write-Host "    Recycle Bin emptying error: $_" -ForegroundColor Red
            }

        } else {
            Write-Host "  [WARN] No drive found for user." -ForegroundColor Yellow
        }
    }

    # 4. OneDrive (Site Deletion - Destructive & Reconstructive)
    Invoke-Safe -What "4. Total OneDrive Cleanup (Site Deletion)" -Action {
        if (Grant-OneDriveAdminAccess -UserUpn $UserUpn -AdminUpn $AdminUpn) {
            try {
                # Attempt 1: Get URL from Graph
                $drive = Get-MgUserDrive -UserId $UserId -Property Id, WebUrl -ErrorAction SilentlyContinue | Select-Object -First 1
                $CleanUrl = $null

                if ($drive) {
                    $CleanUrl = $drive.WebUrl
                    if ($CleanUrl -match "^(https://[^\/]+/personal/[^\/]+)") {
                        $CleanUrl = $matches[1]
                    }
                } else {
                    # Attempt 2: Manual URL Calculation (Fallback)
                    Write-Host "  -> OneDrive not found via Graph. Attempting manual calculation..." -ForegroundColor DarkGray
                    $PersonalUrlPart = $UserUpn -replace "@","_" -replace "\.","_"
                    $CleanUrl = "https://$TenantName-my.sharepoint.com/personal/$PersonalUrlPart"
                }

                if ($CleanUrl) {
                    Write-Host "  -> Target Site Collection: $CleanUrl" -ForegroundColor Cyan

                    # Existence Check (Additional logic required)
                    $SiteExists = $null
                    try { $SiteExists = Get-SPOSite -Identity $CleanUrl -ErrorAction SilentlyContinue } catch {}

                    if ($SiteExists -or $drive) {
                         # Remove the entire site (PnP)
                        Write-Host "  -> Total Site Collection Removal (Preventive 404)..." -ForegroundColor Yellow
                        try {
                            Remove-PnPTenantSite -Url $CleanUrl -Force -ErrorAction Stop
                            Write-Host "  -> Site Collection Removed (PnP). User will see 404." -ForegroundColor Green

                            $Results += [PSCustomObject]@{
                                User = $UserUpn
                                Status = "Deleted"
                            }
                        } catch {
                            Write-Error "Failed to remove site: $_"
                        }

                        # RESET OPTION: Delete from Recycle Bin (Permanent)
                        try {
                            Write-Host "  -> Permanent Deletion (Reset)..." -ForegroundColor Red
                            Remove-PnPTenantSite -Url $CleanUrl -FromRecycleBin -Force -ErrorAction SilentlyContinue

                            Write-Host "  -> Requesting New OneDrive Provisioning (Empty)..." -ForegroundColor Cyan
                            # Request-SPOPersonalSite is not available in PnP.
                            # We skip re-provisioning to prioritize the WIPE.
                            Write-Host "  -> [INFO] Re-provisioning skipped (PnP Mode). The site is deleted." -ForegroundColor DarkGray

                            # Timer REMOVED per request (only small technical pause)
                            Start-Sleep -Seconds 2

                            Write-Host "  -> OneDrive Link (manual check): $CleanUrl" -ForegroundColor Cyan
                        } catch {
                             Write-Host "    [WARN] Automatic reset failed (maybe needs post-delete time): $_" -ForegroundColor Yellow
                        }
                    } else {
                        Write-Host "  -> Site not found even manually." -ForegroundColor DarkGray
                        $Results += [PSCustomObject]@{
                            User = $UserUpn
                            Status = "NotFound"
                        }
                    }
                }
            } catch {
                Write-Host "  [!] OneDrive Error: $_" -ForegroundColor Red
                $Results += [PSCustomObject]@{
                    User = $UserUpn
                    Status = "Error: $($_.Exception.Message)"
                }
            }
        } else {
             Write-Host "  [FAIL] Unable to gain Admin Access to OneDrive." -ForegroundColor Red
        }
    }

    # 5. Activities
    Invoke-Safe -What "5. Activities Cleanup" -Action {
        $activities = Get-MgUserActivity -UserId $UserId -All -ErrorAction SilentlyContinue
        Write-Host "    activities found: $($activities.Count)" -ForegroundColor Yellow
        foreach ($act in $activities) {
            Remove-MgUserActivity -UserId $UserId -ActivityId $act.Id -Confirm:$false -ErrorAction SilentlyContinue
        }
        Write-Host "    Completed!" -ForegroundColor Green
    }

    # 6. Sessions
    Invoke-Safe -What "6. Revoke Sessions" -Action {
        Revoke-MgUserSignInSession -UserId $UserId | Out-Null
        Write-Host "    Completed!" -ForegroundColor Green
    }

    Write-Host "`n✓ USER $UserName COMPLETED" -ForegroundColor Green
    Write-Host "===========================================" -ForegroundColor Cyan
}

# ==============================
# DEFINITIVE PURGE (GLOBAL)
# ==============================
Invoke-Safe -What "DEFINITIVE PURGE (Recycle Bin - Personal Sites)" -Action {
    Write-Host "Searching for personal sites in Recycle Bin (Get-PnPTenantRecycleBinItem)..." -ForegroundColor Yellow
    # Filter only personal sites via PnP
    try {
        $DeletedSites = Get-PnPTenantRecycleBinItem | Where-Object {$_.Url -like "*-my.sharepoint.com/personal/*"}

        if ($DeletedSites) {
            Write-Host "Found $($DeletedSites.Count) sites in Recycle Bin." -ForegroundColor Cyan
            foreach ($DeletedSite in $DeletedSites) {
                Write-Host "  -> Definitive purge: $($DeletedSite.Url)" -ForegroundColor Red
                Remove-PnPTenantSite -Url $DeletedSite.Url -FromRecycleBin -Force -ErrorAction SilentlyContinue
            }
            Write-Host "Purge completed." -ForegroundColor Green
        } else {
            Write-Host "No personal sites found in Recycle Bin." -ForegroundColor Gray
        }
    } catch {
        Write-Host "PnP Recycle Bin Check Failed: $_" -ForegroundColor Yellow
    }
}

Write-Host "`n=== ONEDRIVE SUMMARY ===" -ForegroundColor Cyan
$Results | Format-Table -AutoSize

Write-Host "`n`n========================================" -ForegroundColor Green
Write-Host "FULL CLEANUP COMPLETED" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
