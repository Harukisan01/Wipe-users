# ==============================
# SCRIPT START - LEVEL 3: BANK PROTECTION
# ==============================
param(
    [Parameter(Mandatory = $true)]
    [string]$KeyFilePath,

    [Parameter(Mandatory = $true)]
    [string]$ClientId,

    [Parameter(Mandatory = $true)]
    [string]$ClientSecret,

    [Parameter(Mandatory = $true)]
    [string]$TenantId,

    [Parameter(Mandatory = $false)]
    [string]$UserGroupId
)

$ErrorActionPreference = "Stop"
$RequireTypedConfirmation = $true

# Configuration
if (-not $UserGroupId) {
    # Default Group ID for Testing
    $UserGroupId = "33a31c3c-b300-4879-bc15-6b6aae9c7f6e"
    Write-Host "Using default User Group ID: $UserGroupId" -ForegroundColor Gray
}

# 1. Authenticate with Passkey (Key Vault)
Write-Host "Authenticating with Passkey (Bank Protection Mode)..." -ForegroundColor Cyan

# Check if PasskeyLogin.ps1 exists
if (-not (Test-Path "PasskeyLogin.ps1")) {
    Write-Error "PasskeyLogin.ps1 not found in current directory."
    exit 1
}

# Call PasskeyLogin.ps1
. ./PasskeyLogin.ps1 -KeyFilePath $KeyFilePath `
    -KeyVaultClientId $ClientId `
    -KeyVaultClientSecret $ClientSecret `
    -KeyVaultTenantId $TenantId `
    -PassThru | Out-Null

# Verify Authentication
if (-not $global:ESTSAUTH) {
    Write-Error "Passkey Authentication Failed. No ESTSAUTH cookie found."
    exit 1
}

Write-Host "Passkey Authentication Successful!" -ForegroundColor Green

# 2. Get Microsoft Graph Access Token from Session
function Get-GraphTokenFromSession {
    param($Session)

    Write-Host "Exchanging Session for Graph Access Token..." -ForegroundColor Cyan

    # Ensure Portal Cookies are set (SSO)
    try {
        Write-Host "  -> Visiting Portal Home..." -ForegroundColor Gray
        $null = Invoke-WebRequest -Uri "https://portal.azure.com" -WebSession $Session -Method Get -ErrorAction SilentlyContinue
    } catch {}

    # Hit Portal API to get token
    $TokenUrl = "https://portal.azure.com/api/delegation/token?feature=access_token&scope=user_impersonation&extensionName=Microsoft_Azure_AD"

    try {
        $Response = Invoke-WebRequest -Uri $TokenUrl -WebSession $Session -Method Get -ErrorAction Stop
        $Json = $Response.Content | ConvertFrom-Json
        return $Json.value.access_token
    } catch {
        Write-Warning "Failed to get token via Portal API. Trying alternative..."
        throw "Could not obtain Graph Access Token from Session. Please ensure the user has access to Azure Portal."
    }
}

try {
    $GraphAccessToken = Get-GraphTokenFromSession -Session $global:webSession
    $SecureToken = ConvertTo-SecureString $GraphAccessToken -AsPlainText -Force

    # Connect to Graph
    Connect-MgGraph -AccessToken $SecureToken -NoWelcome
    $ctx = Get-MgContext
    Write-Host "Connected to Microsoft Graph: $($ctx.Account)" -ForegroundColor Green
} catch {
    Write-Error "Failed to connect to Graph with Passkey session: $_"
    Write-Host "Fallback: Please run 'Connect-MgGraph' manually if needed, or ensure the Passkey user has Portal access." -ForegroundColor Yellow
    exit 1
}

# 3. SharePoint Online Authentication
# Determine Admin URL to get Tenant Name
try {
    $RootSite = Get-MgSite -Filter "siteCollection/root ne null" -Select "webUrl" -ErrorAction Stop
    $TenantHost = ([Uri]$RootSite.WebUrl).Host
    $TenantName = $TenantHost -replace "\.sharepoint\.com", ""
    $AdminUrl = "https://$TenantName-admin.sharepoint.com"
} catch {
    # Fallback: Try to guess from default domain (often unreliable but better than prompt in automation)
    try {
        $Org = Get-MgOrganization
        $OnMicrosoftDomain = $Org.VerifiedDomains | Where-Object { $_.Name -like "*.onmicrosoft.com" } | Select-Object -First 1 -ExpandProperty Name
        $TenantName = $OnMicrosoftDomain -replace "\.onmicrosoft\.com", ""
        $AdminUrl = "https://$TenantName-admin.sharepoint.com"
    } catch {
        Write-Error "Could not detect SharePoint Admin URL automatically. Please ensure the user has access to Graph or provide the URL via parameter."
        exit 1
    }
}

Write-Host "Connecting to SharePoint Online..." -ForegroundColor Cyan
try {
    Write-Warning "Passkey auth for SharePoint PowerShell is limited. You may be prompted to sign in again for SharePoint."
    Connect-SPOService -Url $AdminUrl
    Write-Host "Connected to SharePoint!" -ForegroundColor Green
} catch {
    Write-Error "SharePoint Connection Failed: $_"
    Write-Error "CRITICAL: SharePoint connection failed. Exiting to prevent partial wipe."
    exit 1
}

# ==============================
# WIPE LOGIC (Robust)
# ==============================

$DryRun = $false

function Confirm-DestructiveAction {
    param([string]$Title, [string]$Details)
    Write-Host ""
    Write-Host "=== $Title ===" -ForegroundColor Yellow
    Write-Host $Details -ForegroundColor Yellow
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

function Clean-OneDriveFolder {
    param($UserId, $DriveId, $FolderPath)
    try {
        $TargetItem = Get-MgUserDriveItem -UserId $UserId -DriveId $DriveId -Path $FolderPath -ErrorAction SilentlyContinue
        if ($TargetItem) {
            Write-Host "  -> Folder '$FolderPath' found"
             Remove-MgUserDriveItem -UserId $UserId -DriveId $DriveId -DriveItemId $TargetItem.Id -Confirm:$false -ErrorAction SilentlyContinue
            Write-Host "    Deleted." -ForegroundColor Green
        }
    } catch {
        Write-Host "  -> Cleanup Error '$FolderPath': $_" -ForegroundColor Red
    }
}

Write-Host "`nRetrieving members of group $UserGroupId..." -ForegroundColor Cyan
$Users = Get-MgGroupMember -GroupId $UserGroupId -All -ErrorAction SilentlyContinue | Where-Object { $_.AdditionalProperties.'@odata.type' -eq "#microsoft.graph.user" }

if (-not $Users) {
    Write-Host "No users found." -ForegroundColor Yellow
    exit
}

Write-Host "Users found: $($Users.Count)" -ForegroundColor Green
$ok = Confirm-DestructiveAction "LEVEL 3 WIPE" "Users: $($Users.Count). This will wipe Mailboxes and OneDrive."
if (-not $ok) { exit }

foreach ($UserRef in $Users) {
    $UserId = $UserRef.Id
    $User = Get-MgUser -UserId $UserId -Property UserPrincipalName,DisplayName -ErrorAction SilentlyContinue
    Write-Host "PROCESSING: $($User.DisplayName)" -ForegroundColor Cyan

    # 1. Email
    Invoke-Safe -What "Email Cleanup" -Action {
        $Messages = Get-MgUserMessage -UserId $UserId -All -Property Id
        foreach ($Msg in $Messages) { Remove-MgUserMessage -UserId $UserId -MessageId $Msg.Id -Confirm:$false }
    }

    # 2. OneDrive Site
    # Using SPO Connection established earlier
    $PersonalUrl = "https://$TenantName-my.sharepoint.com/personal/$($User.UserPrincipalName -replace '[\.@]', '_')"
    Invoke-Safe -What "OneDrive Site Deletion" -Action {
        Remove-SPOSite -Identity $PersonalUrl -NoWait -Confirm:$false -ErrorAction SilentlyContinue
        Remove-SPODeletedSite -Identity $PersonalUrl -NoWait -Confirm:$false -ErrorAction SilentlyContinue
        Request-SPOPersonalSite -UserEmails $User.UserPrincipalName -NoWait
    }

    # 3. Specific Folders
    Invoke-Safe -What "OneDrive Folders Cleanup" -Action {
        $Drive = Get-MgUserDrive -UserId $UserId -ErrorAction SilentlyContinue
        if ($Drive) {
            Clean-OneDriveFolder -UserId $UserId -DriveId $Drive.Id -FolderPath "/shared"
            Clean-OneDriveFolder -UserId $UserId -DriveId $Drive.Id -FolderPath "/favorites"
            Clean-OneDriveFolder -UserId $UserId -DriveId $Drive.Id -FolderPath "/my"
        }
    }

    # 4. Activities & Sessions
    Invoke-Safe -What "Activities & Sessions" -Action {
        $activities = Get-MgUserActivity -UserId $UserId -All -ErrorAction SilentlyContinue
        foreach ($act in $activities) { Remove-MgUserActivity -UserId $UserId -ActivityId $act.Id -Confirm:$false }
        Revoke-MgUserSignInSession -UserId $UserId | Out-Null
    }
}

Write-Host "Done."
