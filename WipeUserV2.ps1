# ==============================
# INIZIO SCRIPT
# ==============================
$ErrorActionPreference = "Stop"

# Converti SecureString a Plain Text per le richieste HTTP
$PlainSecret = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($ClientSecret))

# Funzione Helper per ottenere Token OAuth2
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
        Write-Error "Impossibile ottenere il token per $Resource : $_"
        throw
    }
}

Write-Host "WipeUser v2.0 - Modalità App Registration" -ForegroundColor Cyan
Write-Host "Tenant: $TenantId | ClientId: $ClientId" -ForegroundColor Gray

# 1. Autenticazione Microsoft Graph
Write-Host "`nRichiesta Token Microsoft Graph..." -ForegroundColor Cyan
try {
    $GraphToken = Get-OAuthToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $PlainSecret -Resource "https://graph.microsoft.com"
    $SecureGraphToken = ConvertTo-SecureString $GraphToken -AsPlainText -Force
    
    # Disconnetti eventuali sessioni precedenti
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    
    # Connetti con Access Token
    Connect-MgGraph -AccessToken $SecureGraphToken -NoWelcome
    $ctx = Get-MgContext
    Write-Host "Connesso a Microsoft Graph." -ForegroundColor Green
} catch {
    Write-Error "Errore connessione Graph: $_"
    exit 1
}

# 2. Autenticazione SharePoint Online
try {
    Write-Host "`nConfigurazione SharePoint Online..." -ForegroundColor Cyan

    # Installa modulo se mancante
    if (-not (Get-Module -ListAvailable -Name Microsoft.Online.SharePoint.PowerShell)) {
        Write-Warning "Modulo SharePoint non trovato. Installazione..."
        Install-Module -Name Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module Microsoft.Online.SharePoint.PowerShell -WarningAction SilentlyContinue -ErrorAction Stop

    # Determina URL Admin
    Write-Host "Rilevamento URL Admin..." -ForegroundColor Gray
    try {
        $RootSite = Get-MgSite -Filter "siteCollection/root ne null" -Select "webUrl" -ErrorAction Stop
        if ($RootSite -and $RootSite.WebUrl) {
            $TenantHost = ([Uri]$RootSite.WebUrl).Host
            $TenantName = $TenantHost -replace "\.sharepoint\.com", ""
            $AdminUrl = "https://$TenantName-admin.sharepoint.com"
            Write-Host "  -> Admin URL: $AdminUrl" -ForegroundColor DarkGray
        } else {
            throw "Root Site non trovato."
        }
    } catch {
        # Fallback legacy
        Write-Warning "Metodo Graph fallito. Riprovo metodo legacy..."
        $Org = Get-MgOrganization
        $OnMicrosoftDomain = $Org.VerifiedDomains | Where-Object { $_.Name -like "*.onmicrosoft.com" } | Select-Object -First 1 -ExpandProperty Name
        $TenantName = $OnMicrosoftDomain -replace "\.onmicrosoft\.com", ""
        $AdminUrl = "https://$TenantName-admin.sharepoint.com"
        Write-Host "  -> Admin URL (Legacy): $AdminUrl" -ForegroundColor DarkGray
    }

    # Richiedi Token SPO Admin
    Write-Host "Richiesta Token SharePoint Admin..." -ForegroundColor Cyan
    # Lo scope per SPO Admin è solitamente l'URL admin + /.default
    $SpoToken = Get-OAuthToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $PlainSecret -Resource $AdminUrl
    
    # TENTATIVO 1: Connect-SPOService (modern auth, se supportata)
    try {
        # Alcune versioni interne supportano -AccessToken, ma per sicurezza usiamo PnP se il modulo standard fallisce o non ha il parametro.
        # Verifichiamo se Connect-SPOService ha -AccessToken
        if (Get-Command Connect-SPOService | Select-Object -ExpandProperty Parameters | Where-Object {$_.Key -eq "AccessToken"}) {
             Connect-SPOService -Url $AdminUrl -AccessToken $SpoToken -ErrorAction Stop
             Write-Host "Connesso a SharePoint Online (Native Module)." -ForegroundColor Green
        } else {
             # Fallback: Usiamo PnP PowerShell se installato, o cerchiamo di usare Graph per le operazioni SPO.
             # PnP è lo standard di fatto per App-Only moderno.
             Write-Warning "Il modulo SharePoint installato non supporta -AccessToken."
             Write-Warning "Si consiglia di installare PnP.PowerShell per l'App-Only Auth completa."
             
             throw "Modulo SPO obsoleto o non supportato per App-Only senza certificato."
        }
    } catch {
         Write-Warning "Connessione SPO fallita: $_"
         Write-Host "NOTA: Per usare App Registration con SPO, assicurati di avere l'ultima versione di Microsoft.Online.SharePoint.PowerShell" -ForegroundColor Yellow
    }

} catch {
    Write-Error "Errore Inizializzazione SharePoint: $_"
}

# ==============================
# FUNZIONI UTILI
# ==============================

function Confirm-DestructiveAction {
    param([string]$Title, [string]$Details)

    Write-Host ""
    Write-Host "=== $Title ===" -ForegroundColor Yellow
    Write-Host $Details -ForegroundColor Yellow
    Write-Host "DryRun=$DryRun" -ForegroundColor Yellow
    Write-Host ""

    if (-not $RequireTypedConfirmation) { return $true }
    $typed = Read-Host "Scrivi 'ESEGUI' per confermare"
    return ($typed -eq "ESEGUI")
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

# Funzione modificata per App-Only
function Grant-OneDriveAdminAccess {
    param([string]$UserUpn, [string]$AdminUpn) # AdminUpn qui è il ServicePrincipal ID o vuoto se usiamo App Permissions

    # Con App-Only (Sites.FullControl.All), l'app ha già accesso a tutto.
    # Non è necessario aggiungersi esplicitamente come SiteCollectionAdmin se usiamo Graph.
    # Tuttavia, se usiamo comandi SPO legacy, potrebbe servire.
    # Per ora restituiamo true assumendo che l'App abbia i permessi.
    return $true
}

function Remove-DriveItemRecursively {
    param($UserId, $DriveId, $FolderId)
    # Stessa logica di WipeUser.ps1 (Graph)
    $Items = Get-MgUserDriveItem -UserId $UserId -DriveId $DriveId -ParentId $FolderId -All -ErrorAction SilentlyContinue
    if ($Items) {
        foreach ($Item in $Items) {
            if ($Item.Folder -ne $null) {
                Remove-DriveItemRecursively -UserId $UserId -DriveId $DriveId -FolderId $Item.Id
                Write-Host "    -> Eliminazione cartella interna: $($Item.Name)"
                Remove-MgUserDriveItem -UserId $UserId -DriveId $DriveId -DriveItemId $Item.Id -Confirm:$false -ErrorAction SilentlyContinue
            } else {
                Write-Host "    -> Eliminazione file interno: $($Item.Name)"
                Remove-MgUserDriveItem -UserId $UserId -DriveId $DriveId -DriveItemId $Item.Id -Confirm:$false -ErrorAction SilentlyContinue
            }
        }
    }
}

function Clean-OneDriveFolder {
    param($UserId, $DriveId, $FolderPath)
    # Stessa logica di WipeUser.ps1 (Graph)
    try {
        $TargetItem = Get-MgUserDriveItem -UserId $UserId -DriveId $DriveId -Path $FolderPath -ErrorAction SilentlyContinue
        if ($TargetItem) {
            Write-Host "  -> Cartella '$FolderPath' trovata (ID: $($TargetItem.Id))"
            Remove-DriveItemRecursively -UserId $UserId -DriveId $DriveId -FolderId $TargetItem.Id
            if ($FolderPath -ne "/" -and $FolderPath -ne "") {
                Write-Host "    -> Eliminazione cartella radice '$FolderPath'"
                Remove-MgUserDriveItem -UserId $UserId -DriveId $DriveId -DriveItemId $TargetItem.Id -Confirm:$false -ErrorAction SilentlyContinue
            }
            Write-Host "    Completato!" -ForegroundColor Green
        } else {
            Write-Host "  -> Cartella '$FolderPath' non trovata." -ForegroundColor Gray
        }
    } catch {
        Write-Host "  -> Errore pulizia '$FolderPath': $_" -ForegroundColor Red
    }
}

# ==============================
# LOGICA PRINCIPALE
# ==============================

Write-Host "`nRecupero membri gruppo $UserGroupId..." -ForegroundColor Cyan
$Users = Get-MgGroupMember -GroupId $UserGroupId -All -ErrorAction SilentlyContinue |
    Where-Object { $_.AdditionalProperties.'@odata.type' -eq "#microsoft.graph.user" }

if (-not $Users) {
    Write-Host "Nessun utente trovato nel gruppo." -ForegroundColor Yellow
    exit
}

Write-Host "Utenti trovati: $($Users.Count)" -ForegroundColor Green

$okUsers = Confirm-DestructiveAction -Title "PULIZIA GENERALE & ONEDRIVE (DISTRUTTIVA)" -Details "Utenti: $($Users.Count). Operazioni: Mailbox, Posta Eliminata, Cartelle (Shared, Favorites, My), Recycle Bin, OneDrive (Reset), Attività, Sessioni."

if (-not $okUsers) {
    Write-Host "Annullato."
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
    Invoke-Safe -What "1. Pulizia Email" -Action {
        $Messages = Get-MgUserMessage -UserId $UserId -All -Property Id -ErrorAction SilentlyContinue
        Write-Host "    messaggi trovati: $($Messages.Count)" -ForegroundColor Yellow
        foreach ($Msg in $Messages) {
            Remove-MgUserMessage -UserId $UserId -MessageId $Msg.Id -Confirm:$false
        }
        Write-Host "    Completato!" -ForegroundColor Green
    }

    # 2. Posta Eliminata (Graph)
    Invoke-Safe -What "2. Pulizia Posta Eliminata" -Action {
        $Deleted = Get-MgUserMailFolder -UserId $UserId -All -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName -eq "Deleted Items" } | Select-Object -First 1
        if ($Deleted) {
            $DeletedMessages = Get-MgUserMailFolderMessage -UserId $UserId -MailFolderId $Deleted.Id -All -Property Id
            Write-Host "    messaggi eliminati trovati: $($DeletedMessages.Count)" -ForegroundColor Yellow
            foreach ($Msg in $DeletedMessages) {
                Remove-MgUserMailFolderMessage -UserId $UserId -MailFolderId $Deleted.Id -MessageId $Msg.Id -Confirm:$false
            }
            Write-Host "    Completato!" -ForegroundColor Green
        }
    }

    # 3. Cartelle Specifiche OneDrive (Graph)
    Invoke-Safe -What "3. Pulizia Cartelle Specifiche OneDrive" -Action {
        $Drive = Get-MgUserDrive -UserId $UserId -ErrorAction SilentlyContinue
        if ($Drive) {
            Clean-OneDriveFolder -UserId $UserId -DriveId $Drive.Id -FolderPath "/shared"
            Clean-OneDriveFolder -UserId $UserId -DriveId $Drive.Id -FolderPath "/favorites"
            Clean-OneDriveFolder -UserId $UserId -DriveId $Drive.Id -FolderPath "/my"

            Write-Host "  -> Svuotamento Cestino OneDrive..."
            try {
                $DeletedItems = Get-MgUserDriveItem -UserId $UserId -DriveId $Drive.Id -Filter "deleted ne null" -All -ErrorAction SilentlyContinue
                if ($DeletedItems) {
                     Write-Host "    Elementi nel cestino: $($DeletedItems.Count)"
                     foreach ($DelItem in $DeletedItems) {
                         Remove-MgUserDriveItem -UserId $UserId -DriveId $Drive.Id -DriveItemId $DelItem.Id -Confirm:$false -ErrorAction SilentlyContinue
                     }
                     Write-Host "    Cestino svuotato." -ForegroundColor Green
                } else {
                     Write-Host "    Cestino vuoto o inaccessibile via Graph." -ForegroundColor Gray
                }
            } catch {
                Write-Host "    Errore svuotamento cestino: $_" -ForegroundColor Red
            }

        } else {
            Write-Host "  [WARN] Nessun drive trovato per l'utente." -ForegroundColor Yellow
        }
    }

    # 4. OneDrive (Site Deletion - SPO)
    Invoke-Safe -What "4. Pulizia Totale OneDrive (Site Deletion)" -Action {
        # App-Only ha già accesso Admin
        if (Grant-OneDriveAdminAccess -UserUpn $UserUpn -AdminUpn $ClientId) {
            try {
                # Tentativo 1: Ottieni URL da Graph
                $drive = Get-MgUserDrive -UserId $UserId -Property Id, WebUrl -ErrorAction SilentlyContinue | Select-Object -First 1
                $CleanUrl = $null

                if ($drive) {
                    $CleanUrl = $drive.WebUrl
                    if ($CleanUrl -match "^(https://[^\/]+/personal/[^\/]+)") {
                        $CleanUrl = $matches[1]
                    }
                } else {
                    Write-Host "  -> OneDrive non trovato via Graph. Tentativo calcolo manuale..." -ForegroundColor DarkGray
                    $PersonalUrlPart = $UserUpn -replace "@","_" -replace "\.","_"
                    $CleanUrl = "https://$TenantName-my.sharepoint.com/personal/$PersonalUrlPart"
                }

                if ($CleanUrl) {
                    Write-Host "  -> Target Site Collection: $CleanUrl" -ForegroundColor Cyan

                    # Verifica esistenza (Logica aggiuntiva richiesta)
                    $SiteExists = $null
                    try { $SiteExists = Get-SPOSite -Identity $CleanUrl -ErrorAction SilentlyContinue } catch {}

                    if ($SiteExists -or $drive) {
                         # Rimuovi il sito intero
                        Write-Host "  -> Rimozione Totale Site Collection (Preventivo 404)..." -ForegroundColor Yellow
                        try {
                            Remove-SPOSite -Identity $CleanUrl -NoWait -Confirm:$false -ErrorAction Stop
                            Write-Host "  -> Site Collection Rimossa. L'utente vedrà 404." -ForegroundColor Green
                            $Results += [PSCustomObject]@{ User = $UserUpn; Status = "Deleted" }
                        } catch {
                            Write-Error "Errore Remove-SPOSite: $_"
                            $Results += [PSCustomObject]@{ User = $UserUpn; Status = "Error Delete" }
                        }

                        # OPZIONE RESET: Elimina dal Cestino (Permanente) e Ricrea (Vuoto)
                        try {
                            Write-Host "  -> Eliminazione Definitiva (Reset)..." -ForegroundColor Red
                            Remove-SPODeletedSite -Identity $CleanUrl -NoWait -Confirm:$false -ErrorAction SilentlyContinue

                            Write-Host "  -> Richiesta provisioning Nuovo OneDrive (Vuoto)..." -ForegroundColor Cyan
                            Request-SPOPersonalSite -UserEmails $UserUpn -NoWait -ErrorAction Stop
                            Write-Host "  -> OK. Il nuovo sito sarà pronto a breve." -ForegroundColor Green

                            Write-Host "  -> Link OneDrive (verifica manuale): $CleanUrl" -ForegroundColor Cyan
                        } catch {
                             Write-Host "    [WARN] Reset automatico non riuscito: $_" -ForegroundColor Yellow
                        }
                    } else {
                        Write-Host "  -> Site non trovato neanche manualmente." -ForegroundColor DarkGray
                        $Results += [PSCustomObject]@{ User = $UserUpn; Status = "NotFound" }
                    }
                }
            } catch {
                Write-Host "  [!] Errore OneDrive: $_" -ForegroundColor Red
                $Results += [PSCustomObject]@{ User = $UserUpn; Status = "Error: $($_.Exception.Message)" }
            }
        }
    }

    # 5. Attività (Graph)
    Invoke-Safe -What "5. Pulizia Attività" -Action {
        $activities = Get-MgUserActivity -UserId $UserId -All -ErrorAction SilentlyContinue
        Write-Host "    attività trovate: $($activities.Count)" -ForegroundColor Yellow
        foreach ($act in $activities) {
            Remove-MgUserActivity -UserId $UserId -ActivityId $act.Id -Confirm:$false -ErrorAction SilentlyContinue
        }
        Write-Host "    Completato!" -ForegroundColor Green
    }

    # 6. Sessioni (Graph)
    Invoke-Safe -What "6. Revoca Sessioni" -Action {
        Revoke-MgUserSignInSession -UserId $UserId | Out-Null
        Write-Host "    Completato!" -ForegroundColor Green
    }

    Write-Host "`n✓ UTENTE $UserName COMPLETATO" -ForegroundColor Green
    Write-Host "===========================================" -ForegroundColor Cyan
}

# ==============================
# PURGE DEFINITIVO (GLOBALE)
# ==============================
Invoke-Safe -What "PURGE DEFINITIVO (Recycle Bin - Personal Sites)" -Action {
    Write-Host "Ricerca siti personali nel cestino (Get-SPODeletedSite)..." -ForegroundColor Yellow
    $DeletedSites = Get-SPODeletedSite | Where-Object {$_.Url -like "*-my.sharepoint.com/personal/*"}

    if ($DeletedSites) {
        Write-Host "Trovati $($DeletedSites.Count) siti nel cestino." -ForegroundColor Cyan
        foreach ($DeletedSite in $DeletedSites) {
            Write-Host "  -> Purge definitivo: $($DeletedSite.Url)" -ForegroundColor Red
            Remove-SPODeletedSite -Identity $DeletedSite.Url -Confirm:$false -ErrorAction SilentlyContinue
        }
        Write-Host "Purge completato." -ForegroundColor Green
    } else {
        Write-Host "Nessun sito personale trovato nel cestino." -ForegroundColor Gray
    }
}

Write-Host "`n=== RIEPILOGO ONEDRIVE ===" -ForegroundColor Cyan
$Results | Format-Table -AutoSize

Write-Host "`n`n========================================" -ForegroundColor Green
Write-Host "FINE PULIZIA COMPLETA (APP REGISTRATION)" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
