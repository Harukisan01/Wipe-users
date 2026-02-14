# ==============================
# SCRIPT START - LEVEL 2: AUTOMATION (PURE PNP)
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

# ==============================
# 1. PnP Module Verification
# ==============================
Write-Host "Verifying PnP.PowerShell module..." -ForegroundColor Cyan
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Warning "PnP.PowerShell module not found. Installing..."
    Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
}
Import-Module PnP.PowerShell -WarningAction SilentlyContinue -ErrorAction Stop

# ==============================
# 2. Authentication & Admin URL Discovery
# ==============================
Write-Host "Connecting to Microsoft 365 (PnP App-Only)..." -ForegroundColor Cyan

try {
    # Initial connection to get Graph access (using the 'common' or tenant-specific endpoint)
    # We use the Tenant ID in the URL to ensure correct context
    $GraphUrl = "https://graph.microsoft.com"

    # We must connect to a site to use PnP fully, but for Graph operations we can use Connect-PnPOnline with ClientId/Secret
    # To discover the Admin URL, we first connect to the Tenant Root (e.g. contoso.sharepoint.com) if we can guess it,
    # OR we use Graph to find it.

    # Let's connect to Graph using PnP first to discover details
    Connect-PnPOnline -ClientId $ClientId -ClientSecret $PlainSecret -Tenant $TenantId -Url "https://$TenantId" -ErrorAction SilentlyContinue
    # Note: Connecting to "https://TenantID" isn't a valid SPO URL usually, but PnP might allow Graph operations.
    # A safer bet is constructing the admin URL if possible, or using Graph directly.

    # Let's try to infer from TenantId? No, UUID doesn't help.
    # We need the onmicrosoft domain.
    # BUT, PnP needs a URL to connect to SPO.

    # Solution: We can't easily discover the Admin URL via PnP App-Only if we don't know the domain *unless* we have a known site.
    # Assumption: The user provides the Admin URL or we accept a standard "TenantName" variable.
    # Fallback: Use the Tenant ID to get Verified Domains via Graph REST?
    # PnP `Connect-PnPOnline` requires a URL for SPO.

    # Workaround: Use basic Invoke-RestMethod for the *first* discovery step to find the tenant domain.
    # Get Token for Graph
    $TokenUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $Body = @{
        grant_type    = "client_credentials"
        client_id     = $ClientId
        client_secret = $PlainSecret
        scope         = "https://graph.microsoft.com/.default"
    }
    $TokenResp = Invoke-RestMethod -Uri $TokenUrl -Method Post -Body $Body
    $GraphToken = $TokenResp.access_token

    # Get Tenant Info
    $OrgHeaders = @{ Authorization = "Bearer $GraphToken" }
    $OrgData = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/organization" -Headers $OrgHeaders
    $VerifiedDomains = $OrgData.value[0].verifiedDomains
    $OnMicrosoftDomain = $VerifiedDomains | Where-Object { $_.name -like "*.onmicrosoft.com" } | Select-Object -First 1 -ExpandProperty name
    $TenantName = $OnMicrosoftDomain -replace "\.onmicrosoft\.com", ""
    $AdminUrl = "https://$TenantName-admin.sharepoint.com"

    Write-Host "  -> Detected Admin URL: $AdminUrl" -ForegroundColor DarkGray

    # Now Connect to PnP properly targeting the Admin Site
    Connect-PnPOnline -Url $AdminUrl -ClientId $ClientId -ClientSecret $PlainSecret -ErrorAction Stop
    Write-Host "Connected to SharePoint Admin (PnP)!" -ForegroundColor Green

} catch {
    Write-Error "Initialization Failed: $_"
    exit 1
}

# ==============================
# HELPER FUNCTIONS
# ==============================

function Send-NotificationEmail {
    param($To, $Subject, $Body)
    if (-not $To) { return }
    Write-Host "Sending notification email to $To..." -ForegroundColor Cyan

    try {
        # Using PnP Graph for Mail
        # Note: Sending mail via Graph App-Only requires 'Mail.Send' Application permission.
        # Format for Send-PnPMail doesn't strictly exist as a cmdlet for arbitrary messages in older PnP?
        # New PnP has Send-PnPMail but it uses specific schemas.
        # We use Invoke-PnPGraphMethod.

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
        } # No conversion to JSON needed for Invoke-PnPGraphMethod if passed as object?
        # Actually Invoke-PnPGraphMethod expects Content to be an object or JSON string.

        Invoke-PnPGraphMethod -Url "v1.0/users/$To/sendMail" -Method Post -Content $EmailJson
        Write-Host "Email sent successfully!" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to send email: $_"
    }
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
    $Response = Invoke-PnPGraphMethod -Url "v1.0/groups/$UserGroupId/members?`$select=id,userPrincipalName,displayName" -Method Get
    $Users = $Response.value
} catch {
    Write-Error "Failed to retrieve group members: $_"
    exit 1
}

if (-not $Users) {
    Write-Host "No users found in the group." -ForegroundColor Yellow
    exit
}

Write-Host "Users found: $($Users.Count)" -ForegroundColor Green

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
        $Msgs = (Invoke-PnPGraphMethod -Url "v1.0/users/$UserId/messages?`$select=id" -Method Get).value
        foreach ($Msg in $Msgs) {
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
    try {
        $DeletedSites = Get-PnPTenantRecycleBinItem | Where-Object {$_.Url -like "*-my.sharepoint.com/personal/*"}
        foreach ($DeletedSite in $DeletedSites) {
            Write-Host "  -> Purge: $($DeletedSite.Url)"
            Remove-PnPTenantSite -Url $DeletedSite.Url -FromRecycleBin -Force -ErrorAction SilentlyContinue
        }
    } catch {}
}

# Send Email Notification
if ($NotificationEmail) {
    $EmailSubject = "Wipe User Automation Report - $(Get-Date -Format 'yyyy-MM-dd')"
    $EmailBody = "Wipe User Automation Completed.`n`nProcessed Users: $($Users.Count)`n`nResults:`n"
    if ($Results) { $EmailBody += ($Results | Out-String) }
    Send-NotificationEmail -To $NotificationEmail -Subject $EmailSubject -Body $EmailBody
}

Write-Host "`n`n========================================" -ForegroundColor Green
Write-Host "FULL CLEANUP COMPLETED (PURE PNP AUTOMATION)" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
