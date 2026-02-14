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
    $UserGroupId = "33a31c3c-b300-4879-bc15-6b6aae9c7f6e"
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

# 2. SharePoint Online Authentication
try {
    Write-Host "`nConfiguring SharePoint Online..." -ForegroundColor Cyan

    # Install module if missing
    if (-not (Get-Module -ListAvailable -Name Microsoft.Online.SharePoint.PowerShell)) {
        Write-Warning "SharePoint module not found. Installing..."
        Install-Module -Name Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module Microsoft.Online.SharePoint.PowerShell -WarningAction SilentlyContinue -ErrorAction Stop

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
        Write-Warning "Graph method failed. Retrying legacy method..."
        $Org = Get-MgOrganization
        $OnMicrosoftDomain = $Org.VerifiedDomains | Where-Object { $_.Name -like "*.onmicrosoft.com" } | Select-Object -First 1 -ExpandProperty Name
        $TenantName = $OnMicrosoftDomain -replace "\.onmicrosoft\.com", ""
        $AdminUrl = "https://$TenantName-admin.sharepoint.com"
        Write-Host "  -> Admin URL (Legacy): $AdminUrl" -ForegroundColor DarkGray
    }

    # Request SPO Admin Token
    Write-Host "Requesting SharePoint Admin Token..." -ForegroundColor Cyan
    # Scope for SPO Admin is usually admin URL + /.default
    $SpoToken = Get-OAuthToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $PlainSecret -Resource $AdminUrl

    # ATTEMPT 1: Connect-SPOService (modern auth, if supported)
    try {
        # Check if Connect-SPOService has -AccessToken
        if (Get-Command Connect-SPOService | Select-Object -ExpandProperty Parameters | Where-Object {$_.Key -eq "AccessToken"}) {
             Connect-SPOService -Url $AdminUrl -AccessToken $SpoToken -ErrorAction Stop
             Write-Host "Connected to SharePoint Online (Native Module)." -ForegroundColor Green
        } else {
             # Fallback: Suggest PnP or similar
             Write-Warning "Installed SharePoint module does not support -AccessToken."
             Write-Warning "It is recommended to install PnP.PowerShell for full App-Only Auth."

             throw "SPO Module obsolete or unsupported for App-Only without certificate."
        }
    } catch {
         Write-Warning "SPO Connection Failed: $_"
         Write-Host "NOTE: To use App Registration with SPO, ensure you have the latest Microsoft.Online.SharePoint.PowerShell" -ForegroundColor Yellow
    }

} catch {
    Write-Error "SharePoint Initialization Error: $_"
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

    # 4. OneDrive (Site Deletion - SPO)
    Invoke-Safe -What "4. Total OneDrive Cleanup (Site Deletion)" -Action {
        # App-Only already has Admin access
        if (Grant-OneDriveAdminAccess -UserUpn $UserUpn -AdminUpn $ClientId) {
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

                    # Existence Check
                    $SiteExists = $null
                    try { $SiteExists = Get-SPOSite -Identity $CleanUrl -ErrorAction SilentlyContinue } catch {}

                    if ($SiteExists -or $drive) {
                         # Remove the entire site
                        Write-Host "  -> Total Site Collection Removal (Preventive 404)..." -ForegroundColor Yellow
                        try {
                            Remove-SPOSite -Identity $CleanUrl -NoWait -Confirm:$false -ErrorAction Stop
                            Write-Host "  -> Site Collection Removed. User will see 404." -ForegroundColor Green
                            $Results += [PSCustomObject]@{ User = $UserUpn; Status = "Deleted" }
                        } catch {
                            Write-Error "Remove-SPOSite Error: $_"
                            $Results += [PSCustomObject]@{ User = $UserUpn; Status = "Error Delete" }
                        }

                        # RESET OPTION: Delete from Recycle Bin (Permanent) and Recreate (Empty)
                        try {
                            Write-Host "  -> Permanent Deletion (Reset)..." -ForegroundColor Red
                            Remove-SPODeletedSite -Identity $CleanUrl -NoWait -Confirm:$false -ErrorAction SilentlyContinue

                            Write-Host "  -> Requesting New OneDrive Provisioning (Empty)..." -ForegroundColor Cyan
                            Request-SPOPersonalSite -UserEmails $UserUpn -NoWait -ErrorAction Stop
                            Write-Host "  -> OK. The new site will be ready shortly." -ForegroundColor Green

                            Write-Host "  -> OneDrive Link (manual check): $CleanUrl" -ForegroundColor Cyan
                        } catch {
                             Write-Host "    [WARN] Automatic reset failed: $_" -ForegroundColor Yellow
                        }
                    } else {
                        Write-Host "  -> Site not found even manually." -ForegroundColor DarkGray
                        $Results += [PSCustomObject]@{ User = $UserUpn; Status = "NotFound" }
                    }
                }
            } catch {
                Write-Host "  [!] OneDrive Error: $_" -ForegroundColor Red
                $Results += [PSCustomObject]@{ User = $UserUpn; Status = "Error: $($_.Exception.Message)" }
            }
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
    Write-Host "Searching for personal sites in Recycle Bin (Get-SPODeletedSite)..." -ForegroundColor Yellow
    $DeletedSites = Get-SPODeletedSite | Where-Object {$_.Url -like "*-my.sharepoint.com/personal/*"}

    if ($DeletedSites) {
        Write-Host "Found $($DeletedSites.Count) sites in Recycle Bin." -ForegroundColor Cyan
        foreach ($DeletedSite in $DeletedSites) {
            Write-Host "  -> Definitive purge: $($DeletedSite.Url)" -ForegroundColor Red
            Remove-SPODeletedSite -Identity $DeletedSite.Url -Confirm:$false -ErrorAction SilentlyContinue
        }
        Write-Host "Purge completed." -ForegroundColor Green
    } else {
        Write-Host "No personal sites found in Recycle Bin." -ForegroundColor Gray
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
