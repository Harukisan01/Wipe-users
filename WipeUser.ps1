# ============================== 
# CONFIGURAZIONE 
# ============================== 
# ID del gruppo contenente gli utenti da ripulire 
$UserGroupId = "" 
 
# Opzioni di sicurezza 
$DryRun = $false 
$RequireTypedConfirmation = $true 
 
# ============================== 
# INIZIO SCRIPT - PULIZIA GENERALE & ONEDRIVE 
# ============================== 
$ErrorActionPreference = "Stop" 
 
# Scope necessari per Microsoft Graph 
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
Write-Host "Login a Microsoft Graph..." -ForegroundColor Cyan 
Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null 
Connect-MgGraph -Scopes $Scopes -NoWelcome 
 
$ctx = Get-MgContext 
Write-Host "Connesso Graph: $($ctx.Account) | Tenant: $($ctx.TenantId)" -ForegroundColor Green 
 
# Login SharePoint Online 
try { 
    Write-Host "`nVerifica modulo SharePoint Online..." -ForegroundColor Cyan 
 
    $SPModule = Get-Module -ListAvailable -Name Microsoft.Online.SharePoint.PowerShell | Select-Object -First 1 
    if (-not $SPModule) { 
        Write-Warning "Modulo SharePoint non trovato. Installazione..." 
        Install-Module -Name Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser -Force -AllowClobber 
        $SPModule = Get-Module -ListAvailable -Name Microsoft.Online.SharePoint.PowerShell | Select-Object -First 1 
    } 
 
    if ($SPModule) { 
        Import-Module $SPModule.Path -WarningAction SilentlyContinue -ErrorAction Stop 
    } 
 
    # RECUPERO AVANZATO URL ADMIN 
    try { 
        $RootSite = Get-MgSite -Filter "siteCollection/root ne null" -Select "webUrl" -ErrorAction Stop 
        if ($RootSite -and $RootSite.WebUrl) { 
            $TenantHost = ([Uri]$RootSite.WebUrl).Host 
            $TenantName = $TenantHost -replace "\.sharepoint\.com", "" 
            $AdminUrl = "https://$TenantName-admin.sharepoint.com" 
            Write-Host "  -> URL Admin rilevato da Graph: $AdminUrl" -ForegroundColor DarkGray 
        } else { 
            throw "Impossibile trovare Root Site tramite Graph." 
        } 
    } catch { 
        Write-Warning "Metodo Graph fallito. Riprovo metodo legacy..." 
        $Org = Get-MgOrganization 
        $OnMicrosoftDomain = $Org.VerifiedDomains | Where-Object { $_.Name -like "*.onmicrosoft.com" } | Select-Object -First 1 -ExpandProperty Name 
        $TenantName = $OnMicrosoftDomain -replace "\.onmicrosoft\.com", "" 
        $AdminUrl = "https://$TenantName-admin.sharepoint.com" 
    } 
 
    # Tentativo Connessione SharePoint 
    $connected = $false 
    do { 
        try { 
            Write-Host "Connessione SharePoint ($AdminUrl)..." -ForegroundColor Cyan 
            Connect-SPOService -Url $AdminUrl -ErrorAction Stop 
            Write-Host "Connesso SharePoint!" -ForegroundColor Green 
            $connected = $true 
        } catch { 
            Write-Host "Errore connessione ($AdminUrl): $($_.Exception.Message)" -ForegroundColor Red 
            $userInput = Read-Host "Inserisci URL Admin SharePoint manualmente (es: https://tenant-admin.sharepoint.com)" 
            if (-not [string]::IsNullOrWhiteSpace($userInput)) { 
                $AdminUrl = $userInput.Trim() 
            } else { 
                Write-Error "Nessun URL fornito. Impossibile gestire i permessi." 
                break 
            } 
        } 
    } until ($connected) 
 
} catch { 
    Write-Host "Errore inizializzazione SharePoint: $_" -ForegroundColor Red 
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
 
function Grant-OneDriveAdminAccess { 
    param([string]$UserUpn, [string]$AdminUpn) 
     
    $SanitizedUser = $UserUpn -replace "[\.@]", "_" 
    $Org = Get-MgOrganization 
    $VerifiedObj = $Org.VerifiedDomains | Where-Object { $_.Name -like "*.onmicrosoft.com" } | Select-Object -First 1 
    $TenantPrefix = ($VerifiedObj.Name -split "\.")[0] 
    $OneDriveUrl = "https://$TenantPrefix-my.sharepoint.com/personal/$SanitizedUser" 
     
    try { 
        Set-SPOUser -Site $OneDriveUrl -LoginName $AdminUpn -IsSiteCollectionAdmin $true -ErrorAction Stop 
        return $true 
    } catch { 
        Write-Host "    [WARN] Errore Set-SPOUser ($OneDriveUrl): $_" -ForegroundColor Yellow 
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
    
    try {
        # Cerca l'item nella root (senza slash iniziale per la ricerca by name nella root, o path completo)
        # Nota: Get-MgUserDriveItemByPath usa il path relativo alla root es: /shared
        $TargetItem = Get-MgUserDriveItem -UserId $UserId -DriveId $DriveId -Path $FolderPath -ErrorAction SilentlyContinue
        
        if ($TargetItem) {
            Write-Host "  -> Cartella '$FolderPath' trovata (ID: $($TargetItem.Id))" 
            Remove-DriveItemRecursively -UserId $UserId -DriveId $DriveId -FolderId $TargetItem.Id
            
            # Rimuove la cartella stessa se non è root
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
    Invoke-Safe -What "1. Pulizia Email" -Action { 
        $Messages = Get-MgUserMessage -UserId $UserId -All -Property Id -ErrorAction SilentlyContinue 
        Write-Host "    messaggi trovati: $($Messages.Count)" -ForegroundColor Yellow 
        foreach ($Msg in $Messages) { 
            Remove-MgUserMessage -UserId $UserId -MessageId $Msg.Id -Confirm:$false 
        } 
        Write-Host "    Completato!" -ForegroundColor Green 
    } 
 
    # 2. Posta Eliminata 
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
 
    # 3. Cartelle Specifiche OneDrive (shared, favorites, my, recycle bin)
    Invoke-Safe -What "3. Pulizia Cartelle Specifiche OneDrive" -Action { 
        $Drive = Get-MgUserDrive -UserId $UserId -ErrorAction SilentlyContinue 
        if ($Drive) { 
            # /shared
            Clean-OneDriveFolder -UserId $UserId -DriveId $Drive.Id -FolderPath "/shared"
            
            # /favorites (se esiste come cartella)
            Clean-OneDriveFolder -UserId $UserId -DriveId $Drive.Id -FolderPath "/favorites"

            # /my (se esiste come cartella)
            Clean-OneDriveFolder -UserId $UserId -DriveId $Drive.Id -FolderPath "/my"

            # Recycle Bin (Cestino OneDrive)
            # Nota: Non è una cartella standard, va svuotato via API specifica o iterando gli item cancellati
            Write-Host "  -> Svuotamento Cestino OneDrive..." 
            try {
                # Get deleted items
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
     
    # 4. OneDrive (Site Deletion - Distruttivo & Ricostruttivo) 
    Invoke-Safe -What "4. Pulizia Totale OneDrive (Site Deletion)" -Action { 
        if (Grant-OneDriveAdminAccess -UserUpn $UserUpn -AdminUpn $AdminUpn) { 
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
                    # Tentativo 2: Calcolo manuale URL (Fallback)
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
                        Remove-SPOSite -Identity $CleanUrl -NoWait -Confirm:$false -ErrorAction Stop 
                        Write-Host "  -> Site Collection Rimossa. L'utente vedrà 404." -ForegroundColor Green 
                        
                        $Results += [PSCustomObject]@{
                            User = $UserUpn
                            Status = "Deleted"
                        }

                        # OPZIONE RESET: Elimina dal Cestino (Permanente) e Ricrea (Vuoto) 
                        try { 
                            Write-Host "  -> Eliminazione Definitiva (Reset)..." -ForegroundColor Red 
                            Remove-SPODeletedSite -Identity $CleanUrl -NoWait -Confirm:$false -ErrorAction SilentlyContinue  
                             
                            Write-Host "  -> Richiesta provisioning Nuovo OneDrive (Vuoto)..." -ForegroundColor Cyan 
                            Request-SPOPersonalSite -UserEmails $UserUpn -NoWait -ErrorAction Stop 
                            Write-Host "  -> OK. Il nuovo sito sarà pronto a breve (15-60 min)." -ForegroundColor Green 
                             
                            # Timer RIMOSSO su richiesta (solo piccola pausa tecnica)
                            Start-Sleep -Seconds 2
                            
                            Write-Host "  -> Link OneDrive (verifica manuale): $CleanUrl" -ForegroundColor Cyan 
                        } catch { 
                             Write-Host "    [WARN] Reset automatico non riuscito (forse serve tempo post-delete): $_" -ForegroundColor Yellow 
                        }
                    } else {
                        Write-Host "  -> Site non trovato neanche manualmente." -ForegroundColor DarkGray
                        $Results += [PSCustomObject]@{
                            User = $UserUpn
                            Status = "NotFound"
                        }
                    }
                }
            } catch { 
                Write-Host "  [!] Errore OneDrive: $_" -ForegroundColor Red 
                $Results += [PSCustomObject]@{
                    User = $UserUpn
                    Status = "Error: $($_.Exception.Message)"
                }
            } 
        } else { 
             Write-Host "  [FAIL] Impossibile ottenere Admin Access al OneDrive." -ForegroundColor Red 
        } 
    } 
 
    # 5. Attività 
    Invoke-Safe -What "5. Pulizia Attività" -Action { 
        $activities = Get-MgUserActivity -UserId $UserId -All -ErrorAction SilentlyContinue 
        Write-Host "    attività trovate: $($activities.Count)" -ForegroundColor Yellow 
        foreach ($act in $activities) { 
            Remove-MgUserActivity -UserId $UserId -ActivityId $act.Id -Confirm:$false -ErrorAction SilentlyContinue 
        } 
        Write-Host "    Completato!" -ForegroundColor Green 
    } 
 
    # 6. Sessioni 
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
    # Filtra solo i siti personali per evitare danni collaterali a siti SharePoint aziendali
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
Write-Host "FINE PULIZIA COMPLETA" -ForegroundColor Green 
Write-Host "========================================`n" -ForegroundColor Green
