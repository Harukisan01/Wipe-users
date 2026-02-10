# ==============================
# CONFIGURAZIONE
# ==============================
$UserGroupId = ""

# Safety switches
$DryRun = $false
$RequireTypedConfirmation = $true

# ==============================
# LOGIN INTERATTIVO (Delegated)
# ==============================
$Scopes = @(
    "GroupMember.Read.All",
    "Mail.ReadWrite",
    "Files.ReadWrite.All",
    "User.ReadWrite.All"
)

# Disconnetto per forzare nuovo token/scope
Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null

Write-Host "Login a Microsoft Graph (interattivo)..."
Connect-MgGraph -Scopes $Scopes -NoWelcome

$ctx = Get-MgContext
Write-Host "Connesso come: $($ctx.Account) | Tenant: $($ctx.TenantId)"

# ==============================
# FUNZIONI UTILI
# ==============================
function Test-GuidOrThrow {
    param([string]$Value, [string]$Name)
    $g = [guid]::Empty
    if (-not [guid]::TryParse($Value, [ref]$g)) {
        throw "$Name non è un GUID valido: '$Value'"
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
    $typed = Read-Host "Scrivi ESATTAMENTE 'ESEGUI' per continuare (altro = annulla)"
    return ($typed -eq "ESEGUI")
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

# ==============================
# VALIDAZIONE INPUT
# ==============================
Test-GuidOrThrow -Value $UserGroupId -Name "UserGroupId"

# ==============================
# 0) PREVIEW: elenca membri
# ==============================
Write-Host "`nRecupero membri del gruppo utenti..."
$Users = Get-MgGroupMember -GroupId $UserGroupId -All |
    Where-Object { $_.AdditionalProperties.'@odata.type' -eq "#microsoft.graph.user" }

Write-Host "Utenti trovati: $($Users.Count)"

# ==============================
# 1) OPERAZIONI UTENTI (DISTRUTTIVE)
# ==============================
if ($Users.Count -gt 0) {
    $okUsers = Confirm-DestructiveAction `
        -Title "OPERAZIONI UTENTI (DISTRUTTIVE)" `
        -Details "Per $($Users.Count) utenti:`n- Cancello email (mailbox)`n- Cancello Deleted Items`n- Cancello OneDrive (Root + Cestino)`n- Cancello Attività (Timeline)`n- Revoco sessioni"

    if (-not $okUsers) {
        Write-Host "Operazioni utenti annullate."
    } else {
        foreach ($UserRef in $Users) {
            $UserId = $UserRef.Id
            Write-Host "`n--- Utente: $UserId ---" -ForegroundColor Cyan

            # ---------------- Email ----------------
            Invoke-Safe -What "Elimino messaggi mailbox per $UserId" -Action {
                $Messages = Get-MgUserMessage -UserId $UserId -All -Property Id
                foreach ($Msg in $Messages) {
                    Remove-MgUserMessage -UserId $UserId -MessageId $Msg.Id -Confirm:$false
                }
            }

            # ---------------- Deleted Items ----------------
            Invoke-Safe -What "Elimino 'Deleted Items' per $UserId" -Action {
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

            # ---------------- OneDrive (Root + Recycle Bin) ----------------
            Invoke-Safe -What "Pulisco OneDrive (Root) e Cestino per $UserId" -Action {
                try {
                    $drive = Get-MgUserDrive -UserId $UserId -Property Id -ErrorAction Stop
                } catch {
                    if ($_.Exception.Message -match "Access denied") {
                        Write-Host "  [!] ACCESS DENIED: Non hai accesso al OneDrive di questo utente." -ForegroundColor Red
                        Write-Host "  [?] SOLUZIONE: Aggiungi il tuo account come 'Site Collection Admin'." -ForegroundColor Yellow
                        return
                    } else {
                        Write-Host "  [!] Errore recupero OneDrive: $_" -ForegroundColor Red
                        return
                    }
                }

                if ($drive) {
                    try {
                        $items = Get-MgDriveRootChild -DriveId $drive.Id -All -Property Id, Name -ErrorAction Stop
                        foreach ($item in $items) {
                            try {
                                Remove-MgDriveItem -DriveId $drive.Id -DriveItemId $item.Id -Confirm:$false -ErrorAction Stop
                            } catch {
                                Write-Host "    [!] Errore cancellazione file '$($item.Name)': $_" -ForegroundColor Red
                            }
                        }
                    } catch { Write-Host "  [!] Errore lettura Root: $_" -ForegroundColor Red }

                    try {
                        $binItems = Get-MgDriveRecycleBin -DriveId $drive.Id -All -ErrorAction SilentlyContinue
                        foreach ($binItem in $binItems) {
                            Write-Host "  -> Elimino definitivamente: $($binItem.Name)" -ForegroundColor DarkGray
                            Invoke-MgGraphRequest -Method DELETE -Uri "drives/$($drive.Id)/recycleBin/$($binItem.Id)"
                        }
                    } catch { Write-Host "  [!] Errore accesso Cestino OneDrive: $_" -ForegroundColor Red }
                }
            }

            # ---------------- User Activities (Timeline) ----------------
            Invoke-Safe -What "Elimino Attività Utente (Timeline/Cronologia) per $UserId" -Action {
                try {
                    $activities = Get-MgUserActivity -UserId $UserId -All -ErrorAction SilentlyContinue
                    foreach ($act in $activities) {
                        Remove-MgUserActivity -UserId $UserId -ActivityId $act.Id -Confirm:$false
                    }
                } catch {
                    Write-Host "  [!] Impossibile pulire attività (permessi/feature mancanti): $_" -ForegroundColor DarkGray
                }
            }

            # ---------------- Sessioni ----------------
            Invoke-Safe -What "Revoco sessioni per $UserId" -Action {
                Revoke-MgUserSignInSession -UserId $UserId | Out-Null
            }
        }
    }
} else {
    Write-Host "Nessun utente nel gruppo."
}

Write-Host "`nFINE OPERAZIONI."

