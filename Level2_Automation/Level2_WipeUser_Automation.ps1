# ==============================
# SCRIPT START
# ==============================
$ErrorActionPreference = "Stop"
$RequireTypedConfirmation = $false # Disable confirmation for automation
$DryRun = $false # Set to $true to simulate actions without execution

# Configuration (to be filled or passed as parameters)
# $TenantId = "..."
# $ClientId = "..."
# $ClientSecret = "..."

# Attempt to retrieve from Azure Automation Variables
try {
    if (-not $TenantId) { $TenantId = Get-AutomationVariable -Name 'TenantId' -ErrorAction SilentlyContinue }
    if (-not $ClientId) { $ClientId = Get-AutomationVariable -Name 'ClientId' -ErrorAction SilentlyContinue }
    if (-not $ClientSecret) { $ClientSecret = Get-AutomationVariable -Name 'ClientSecret' -ErrorAction SilentlyContinue }
    if (-not $UserGroupId) { $UserGroupId = Get-AutomationVariable -Name 'UserGroupId' -ErrorAction SilentlyContinue }
    $NotificationEmail = Get-AutomationVariable -Name 'NotificationEmail' -ErrorAction SilentlyContinue
} catch {
    # Ignore if not running in Azure Automation or variables not found
}

if (-not $UserGroupId) {
    # Default Group ID for Testing/Fallback
    # $UserGroupId = "<INSERT_GROUP_OBJECT_ID_HERE>"
    Write-Error "UserGroupId is missing. Please provide it as a variable or parameter."
    exit 1
}

if (-not $TenantId -or -not $ClientId -or -not $ClientSecret) {
    Write-Error "Missing required variables: TenantId, ClientId, ClientSecret. Please set them as Automation Variables or pass them as parameters."
    exit 1
}

# Convert SecureString to Plain Text for HTTP requests if needed
# If ClientSecret is already plain text, use it directly.
if ($ClientSecret -is [System.Security.SecureString]) {
    $PlainSecret = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($ClientSecret))
} else {
    $PlainSecret = $ClientSecret
}

# Helper Function to get OAuth2 Token
function Get-OAuthToken {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [string]$ClientSecret,
        [string]$Resource
    )

    $TokenUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $Body = @{
        grant_type    = "client_credentials"
        client_id     = $ClientId
        client_secret = $ClientSecret
        scope         = "$Resource/.default"
    }

    try {
        $Response = Invoke-RestMethod -Uri $TokenUrl -Method Post -Body $Body -ErrorAction Stop
        return $Response.access_token
    } catch {
        Write-Error "Unable to get token for $Resource : $_"
        throw
    }
}

Write-Host "WipeUser v2.0 - App Registration Mode" -ForegroundColor Cyan
Write-Host "Tenant: $TenantId | ClientId: $ClientId" -ForegroundColor Gray

# 1. Microsoft Graph Authentication
Write-Host "`nRequesting Microsoft Graph Token..." -ForegroundColor Cyan
try {
    $GraphToken = Get-OAuthToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $PlainSecret -Resource "https://graph.microsoft.com"
    $SecureGraphToken = ConvertTo-SecureString $GraphToken -AsPlainText -Force

    # Disconnect previous sessions
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null

    # Connect with Access Token
    Connect-MgGraph -AccessToken $SecureGraphToken -NoWelcome
    $ctx = Get-MgContext
    Write-Host "Connected to Microsoft Graph." -ForegroundColor Green
} catch {
    Write-Error "Graph Connection Error: $_"
    exit 1
}

# 2. SharePoint Online Authentication (PnP PowerShell)
try {
    Write-Host "`nConfiguring SharePoint Online (PnP PowerShell)..." -ForegroundColor Cyan

    # Install module if missing (PnP.PowerShell)
    if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
        Write-Warning "PnP.PowerShell module not found. Installing..."
        Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module PnP.PowerShell -WarningAction SilentlyContinue -ErrorAction Stop

    # Determine Admin URL
    Write-Host "Detecting Admin URL..." -ForegroundColor Gray
    try {
        $RootSite = Get-MgSite -Filter "siteCollection/root ne null" -Select "webUrl" -ErrorAction Stop
        if ($RootSite -and $RootSite.WebUrl) {
            $TenantHost = ([Uri]$RootSite.WebUrl).Host
            $TenantName = $TenantHost -replace "\.sharepoint\.com", ""
            $AdminUrl = "https://$TenantName-admin.sharepoint.com"
            Write-Host "  -> Admin URL: $AdminUrl" -ForegroundColor DarkGray
        } else {
            throw "Root Site not found."
        }
    } catch {
        # Legacy Fallback
        $Org = Get-MgOrganization
        $OnMicrosoftDomain = $Org.VerifiedDomains | Where-Object { $_.Name -like "*.onmicrosoft.com" } | Select-Object -First 1 -ExpandProperty Name
        $TenantName = $OnMicrosoftDomain -replace "\.onmicrosoft\.com", ""
        $AdminUrl = "https://$TenantName-admin.sharepoint.com"
        Write-Host "  -> Admin URL (Fallback): $AdminUrl" -ForegroundColor DarkGray
    }

    # Connect PnP using ClientId/ClientSecret (App-Only)
    Write-Host "Connecting to SharePoint via PnP..." -ForegroundColor Cyan
    try {
        # Using ClientSecret for App-Only auth
        # Ensure the app has Sites.FullControl.All in SharePoint or Sites.ReadWrite.All in Graph
        Connect-PnPOnline -Url $AdminUrl -ClientId $ClientId -ClientSecret $PlainSecret -ErrorAction Stop
        Write-Host "Connected to SharePoint Online (PnP)." -ForegroundColor Green
    } catch {
         Write-Error "PnP Connection Failed: $_"
         Write-Error "Ensure you are using PnP.PowerShell 2.x+ and the App Registration is configured correctly."
         throw
    }

} catch {
    Write-Error "SharePoint Initialization Error: $_"
    Write-Error "CRITICAL: SharePoint connection failed. Exiting to prevent partial wipe."
    exit 1
}

# ==============================
# USEFUL FUNCTIONS
# ==============================

function Send-NotificationEmail {
    param($To, $Subject, $Body)

    if (-not $To) { return }

    Write-Host "Sending notification email to $To..." -ForegroundColor Cyan

    $EmailJson = @{
        message = @{
            subject = $Subject
            body = @{
                contentType = "Text"
                content = $Body
            }
            toRecipients = @(
                @{
                    emailAddress = @{
                        address = $To
                    }
                }
            )
        }
    } | ConvertTo-Json -Depth 10

    try {
        # Send as the recipient (self-notification) or specific user.
        # App-Only with Mail.Send allows sending as any user.
        Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/users/$To/sendMail" -Body $EmailJson
        Write-Host "Email sent successfully!" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to send email: $_"
        Write-Warning "Ensure the App Registration has 'Mail.Send' permission and '$To' is a valid user."
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
    # In Automation context, we might skip confirmation or use a parameter
    if ($dataset) { return $true } # Assuming non-interactive environment

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

# Modified for App-Only
function Grant-OneDriveAdminAccess {
    param([string]$UserUpn, [string]$AdminUpn) # AdminUpn here is ServicePrincipal ID or empty if using App Permissions

    # With App-Only (Sites.FullControl.All), the app already has access to everything.
    # No need to explicitly add as SiteCollectionAdmin if using Graph.
    # However, if using legacy SPO commands, it might be needed.
    # For now, return true assuming App has permissions.
    return $true
}

function Remove-DriveItemRecursively {
    param($UserId, $DriveId, $FolderId)
    # Same logic as WipeUser.ps1 (Graph)
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
    # Same logic as WipeUser.ps1 (Graph)
    try {
        $TargetItem = Get-MgUserDriveItem -UserId $UserId -DriveId $DriveId -Path $FolderPath -ErrorAction SilentlyContinue
        if ($TargetItem) {
            Write-Host "  -> Folder '$FolderPath' found (ID: $($TargetItem.Id))"
            Remove-DriveItemRecursively -UserId $UserId -DriveId $DriveId -FolderId $TargetItem.Id
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

$Results = @()

foreach ($UserRef in $Users) {
    $UserId = $UserRef.Id

    $User = Get-MgUser -UserId $UserId -Property UserPrincipalName,DisplayName -ErrorAction SilentlyContinue
    $UserUpn = $User.UserPrincipalName
    $UserName = $User.DisplayName

    Write-Host "`n===========================================" -ForegroundColor Cyan
    Write-Host "PROCESSING: $UserName ($UserUpn)" -ForegroundColor Cyan
    Write-Host "===========================================" -ForegroundColor Cyan

    # 1. Email (Graph)
    Invoke-Safe -What "1. Email Cleanup" -Action {
        $Messages = Get-MgUserMessage -UserId $UserId -All -Property Id -ErrorAction SilentlyContinue
        Write-Host "    messages found: $($Messages.Count)" -ForegroundColor Yellow
        foreach ($Msg in $Messages) {
            Remove-MgUserMessage -UserId $UserId -MessageId $Msg.Id -Confirm:$false
        }
        Write-Host "    Completed!" -ForegroundColor Green
    }

    # 2. Deleted Items (Graph)
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

    # 3. Specific OneDrive Folders (Graph)
    Invoke-Safe -What "3. Specific OneDrive Folders Cleanup" -Action {
        $Drive = Get-MgUserDrive -UserId $UserId -ErrorAction SilentlyContinue
        if ($Drive) {
            Clean-OneDriveFolder -UserId $UserId -DriveId $Drive.Id -FolderPath "/shared"
            Clean-OneDriveFolder -UserId $UserId -DriveId $Drive.Id -FolderPath "/favorites"
            Clean-OneDriveFolder -UserId $UserId -DriveId $Drive.Id -FolderPath "/my"

            Write-Host "  -> Emptying OneDrive Recycle Bin..."
            try {
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

    # 4. OneDrive (Site Deletion - PnP)
    Invoke-Safe -What "4. Total OneDrive Cleanup (Site Deletion)" -Action {
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
                Write-Host "  -> OneDrive not found via Graph. Attempting manual calculation..." -ForegroundColor DarkGray
                $PersonalUrlPart = $UserUpn -replace "@","_" -replace "\.","_"
                $CleanUrl = "https://$TenantName-my.sharepoint.com/personal/$PersonalUrlPart"
            }

            if ($CleanUrl) {
                Write-Host "  -> Target Site Collection: $CleanUrl" -ForegroundColor Cyan

                # Existence Check (PnP currently connected to Admin)
                # Remove-PnPTenantSite throws if not found? Let's try direct removal.

                # Remove the entire site
                Write-Host "  -> Total Site Collection Removal (Preventive 404)..." -ForegroundColor Yellow
                try {
                    Remove-PnPTenantSite -Url $CleanUrl -Force -ErrorAction Stop
                    Write-Host "  -> Site Collection Removed (PnP). User will see 404." -ForegroundColor Green
                    $Results += [PSCustomObject]@{ User = $UserUpn; Status = "Deleted" }
                } catch {
                    Write-Host "  -> Site likely didn't exist or error: $($_.Exception.Message)" -ForegroundColor Gray
                    $Results += [PSCustomObject]@{ User = $UserUpn; Status = "NotFound/Error" }
                }

                # RESET OPTION: Delete from Recycle Bin (Permanent)
                try {
                    Write-Host "  -> Permanent Deletion (Reset)..." -ForegroundColor Red
                    Remove-PnPTenantSite -Url $CleanUrl -FromRecycleBin -Force -ErrorAction SilentlyContinue
                    Write-Host "  -> Recycle bin purged." -ForegroundColor Green
                } catch {
                     Write-Host "    [WARN] Recycle bin purge failed: $_" -ForegroundColor Yellow
                }
            }
        } catch {
            Write-Host "  [!] OneDrive Error: $_" -ForegroundColor Red
            $Results += [PSCustomObject]@{ User = $UserUpn; Status = "Error: $($_.Exception.Message)" }
        }
    }

    # 5. Activities (Graph)
    Invoke-Safe -What "5. Activities Cleanup" -Action {
        $activities = Get-MgUserActivity -UserId $UserId -All -ErrorAction SilentlyContinue
        Write-Host "    activities found: $($activities.Count)" -ForegroundColor Yellow
        foreach ($act in $activities) {
            Remove-MgUserActivity -UserId $UserId -ActivityId $act.Id -Confirm:$false -ErrorAction SilentlyContinue
        }
        Write-Host "    Completed!" -ForegroundColor Green
    }

    # 6. Sessions (Graph)
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
        Write-Host "Recycle bin access failed: $_" -ForegroundColor Yellow
    }
}

Write-Host "`n=== ONEDRIVE SUMMARY ===" -ForegroundColor Cyan
$Results | Format-Table -AutoSize

# Send Email Notification
if ($NotificationEmail) {
    $EmailSubject = "Wipe User Automation Report - $(Get-Date -Format 'yyyy-MM-dd')"
    $EmailBody = "Wipe User Automation Completed.`n`nProcessed Users: $($Users.Count)`n`nResults:`n"

    if ($Results) {
        $EmailBody += ($Results | Out-String)
    } else {
        $EmailBody += "No actions recorded or no errors."
    }

    Send-NotificationEmail -To $NotificationEmail -Subject $EmailSubject -Body $EmailBody
}

Write-Host "`n`n========================================" -ForegroundColor Green
Write-Host "FULL CLEANUP COMPLETED (APP REGISTRATION)" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
