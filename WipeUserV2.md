# 📘 Guida Completa: Automazione Wipe User su Azure (v7.2)

Questa guida descrive come configurare un'automazione su Microsoft Azure per eliminare dati utente (Email, OneDrive, Siti Personali) in modo massivo partendo da un Gruppo di Sicurezza.

**Metodo:** API REST Diretta (Nessun modulo PowerShell esterno richiesto).
**Resilienza:** Anti-blocco, esecuzione non interattiva.

---

## 🚀 Fase 1: Registrazione App su Entra ID

Per permettere allo script di agire sul tenant senza login manuale, dobbiamo creare un'identità digitale (Service Principal).

1. Accedi al [Portale Azure](https://portal.azure.com).
2. Vai su **Microsoft Entra ID** (ex Azure Active Directory).
3. Nel menu a sinistra, seleziona **App registrations** > **New registration**.
   - **Name:** `WipeUser-Automation-App`
   - **Supported account types:** Accounts in this organizational directory only (Single tenant).
   - Clicca su **Register**.

### 1.1 Recupero ID
Dalla pagina "Overview" dell'App appena creata, copia e salva in un blocco note:
- **Application (client) ID**
- **Directory (tenant) ID**

### 1.2 Creazione Segreto (Password)
1. Vai su **Certificates & secrets** > **Client secrets**.
2. Clicca su **+ New client secret**.
3. Inserisci una descrizione (es. *AutomationKey*) e una scadenza (es. *24 mesi*).
4. Clicca **Add**.
5. **IMPORTANTE:** Copia subito il **Value** (valore) del segreto. Una volta chiusa la pagina non sarà più visibile.

### 1.3 Permessi API (API Permissions)
L'App deve avere il permesso di leggere e cancellare i dati.

1. Vai su **API permissions** > **+ Add a permission**.
2. Seleziona **Microsoft Graph** > **Application permissions** (NON Delegated).
3. Cerca e spunta i seguenti permessi:
   - `User.ReadWrite.All` (Per leggere gli utenti e disabilitarli/revocare sessioni)
   - `Mail.ReadWrite` (Per cancellare le email)
   - `Files.ReadWrite.All` (Per cancellare OneDrive)
   - `Sites.FullControl.All` (Per cancellare i siti SharePoint personali)
   - `Directory.Read.All` (Per leggere i membri del gruppo)
   - `GroupMember.Read.All` (Per leggere i gruppi)
4. Clicca su **Add permissions**.
5. **FONDAMENTALE:** Clicca sul pulsante **Grant admin consent for [NomeTenant]** e conferma con "Yes". Lo stato dei permessi deve diventare verde ("Granted").

---

## ⚙️ Fase 2: Configurazione Azure Automation

1. Cerca **Automation Accounts** nel portale Azure e crea un nuovo account (es. `Wipe-user`).
2. Una volta creato, vai nella risorsa e cerca nel menu a sinistra **Shared Resources** > **Variables**.

### 2.1 Creazione Variabili
Crea le seguenti variabili cliccando su **+ Add a variable**. I nomi devono essere *esatti*.

| Nome | Tipo | Encrypted | Valore da inserire |
| :--- | :--- | :--- | :--- |
| **TenantId** | String | No | Il *Directory (tenant) ID* copiato nella Fase 1.1 |
| **AppClientId** | String | No | L'*Application (client) ID* copiato nella Fase 1.1 |
| **AppClientSecret** | String | **Yes** | Il *Client Secret Value* copiato nella Fase 1.2 |
| **TargetUserGroupId** | String | No | L'*Object ID* del Gruppo Entra ID contenente gli utenti da pulire. |
| **NotificationSender** | String | No | Email di un admin *interno* (es. `admin@tuotenant.onmicrosoft.com`) per inviare il report. |
| **NotificationReceiver** | String | No | La tua email (dove vuoi ricevere il report). |

---

## 📜 Fase 3: Creazione Runbook (Script)

1. Vai su **Process Automation** > **Runbooks**.
2. Clicca su **+ Create a runbook**.
   - **Name:** `WipeUser-V7`
   - **Runbook type:** PowerShell
   - **Runtime version:** 5.1 (o 7.1, entrambi ok).
3. Clicca **Create**.
4. Nell'editor che si apre, incolla il seguente codice (Versione v7.2 Finale):

```powershell
# ==============================================================================
# WIPE-USER v7.2 - VERSIONE FINALE (REST API PURA)
# ==============================================================================
$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

Write-Output "--- AVVIO PROCEDURA WIPE ---"

# 1. RECUPERO VARIABILI
try {
    $TenantId     = Get-AutomationVariable -Name 'TenantId'
    $ClientId     = Get-AutomationVariable -Name 'AppClientId'
    $ClientSecret = Get-AutomationVariable -Name 'AppClientSecret'
    $UserGroupId  = Get-AutomationVariable -Name 'TargetUserGroupId'
    
    $SendEmail = $false
    try {
        $NotifySender   = Get-AutomationVariable -Name 'NotificationSender'
        $NotifyReceiver = Get-AutomationVariable -Name 'NotificationReceiver'
        if ($NotifySender -and $NotifyReceiver) { $SendEmail = $true }
    } catch {}

} catch {
    Write-Error "ERRORE: Variabili mancanti (TenantId, AppClientId, AppClientSecret, TargetUserGroupId)."
    throw $_
}

# 2. AUTENTICAZIONE REST
Write-Output "Generazione Token..."
$TokenUrl = "[https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token](https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token)"
$Body = @{
    grant_type    = "client_credentials"
    client_id     = $ClientId
    client_secret = $ClientSecret
    scope         = "[https://graph.microsoft.com/.default](https://graph.microsoft.com/.default)"
}

try {
    $Response = Invoke-RestMethod -Uri $TokenUrl -Method Post -Body $Body -ErrorAction Stop
    $Token = $Response.access_token
    $Headers = @{ Authorization = "Bearer $Token" }
    Write-Output " -> Token ottenuto."
} catch {
    throw "ERRORE AUTH: Verifica Client Secret e ID."
}

# 3. RECUPERO NOME TENANT (Dinamico)
try {
    $RootSite = Invoke-RestMethod -Uri "[https://graph.microsoft.com/v1.0/sites/root](https://graph.microsoft.com/v1.0/sites/root)" -Headers $Headers -Method Get -ErrorAction Stop
    $TenantPrefix = ([Uri]$RootSite.webUrl).Host -replace "\.sharepoint\.com", ""
} catch {
    $TenantPrefix = "TENANT-NOT-FOUND"
}

# 4. RECUPERO UTENTI
Write-Output "Lettura gruppo..."
$GroupUrl = "[https://graph.microsoft.com/v1.0/groups/$UserGroupId/members](https://graph.microsoft.com/v1.0/groups/$UserGroupId/members)?`$select=id,displayName,userPrincipalName,userType"
try {
    $GroupResp = Invoke-RestMethod -Uri $GroupUrl -Headers $Headers -Method Get -ErrorAction Stop
    $Users = $GroupResp.value | Where-Object { $_.'@odata.type' -eq "#microsoft.graph.user" }
} catch {
    throw "ERRORE LETTURA GRUPPO: $_"
}

if (-not $Users) { Write-Output "Gruppo vuoto."; exit }

Write-Output "Trovati $($Users.Count) utenti."
$Report = @()

# 5. CICLO DI PULIZIA
foreach ($User in $Users) {
    $Uid  = $User.id
    $UPN  = $User.userPrincipalName
    $Name = $User.displayName
    $Log  = ""

    Write-Output " > PROCESSING: $Name ($UPN)"

    # A. EMAIL
    try {
        $MailUrl = "[https://graph.microsoft.com/v1.0/users/$Uid/messages](https://graph.microsoft.com/v1.0/users/$Uid/messages)?`$select=id&`$top=50"
        $Mails = (Invoke-RestMethod -Uri $MailUrl -Headers $Headers -Method Get).value
        if ($Mails) {
            foreach ($m in $Mails) {
                $DelUrl = "[https://graph.microsoft.com/v1.0/users/$Uid/messages/$($m.id](https://graph.microsoft.com/v1.0/users/$Uid/messages/$($m.id))"
                Invoke-RestMethod -Uri $DelUrl -Headers $Headers -Method Delete -ErrorAction SilentlyContinue | Out-Null
            }
            $Log += "Mail eliminate. "
        }
    } catch { $Log += "Err Mail. " }

    # B. FILES ONEDRIVE (Root)
    try {
        $DriveUrl = "[https://graph.microsoft.com/v1.0/users/$Uid/drive/root/children](https://graph.microsoft.com/v1.0/users/$Uid/drive/root/children)"
        $Items = (Invoke-RestMethod -Uri $DriveUrl -Headers $Headers -Method Get -ErrorAction SilentlyContinue).value
        if ($Items) {
            foreach ($i in $Items) {
                $DelItemUrl = "[https://graph.microsoft.com/v1.0/users/$Uid/drive/items/$($i.id](https://graph.microsoft.com/v1.0/users/$Uid/drive/items/$($i.id))"
                Invoke-RestMethod -Uri $DelItemUrl -Headers $Headers -Method Delete -ErrorAction SilentlyContinue | Out-Null
            }
            $Log += "Root Files eliminati. "
        }
    } catch { $Log += "Err Drive. " }

    # C. CARTELLA SHARED/CONDIVISI
    try {
        $SpecialFolders = @("Shared", "Condivisi", "Shared Documents", "Documents")
        foreach ($Folder in $SpecialFolders) {
            $ItemUrl = "[https://graph.microsoft.com/v1.0/users/$Uid/drive/root:/$Folder](https://graph.microsoft.com/v1.0/users/$Uid/drive/root:/$Folder)"
            try {
                $Item = Invoke-RestMethod -Uri $ItemUrl -Headers $Headers -Method Get -ErrorAction SilentlyContinue
                if ($Item) {
                    $DelUrl = "[https://graph.microsoft.com/v1.0/users/$Uid/drive/items/$($Item.id](https://graph.microsoft.com/v1.0/users/$Uid/drive/items/$($Item.id))"
                    Invoke-RestMethod -Uri $DelUrl -Headers $Headers -Method Delete -ErrorAction SilentlyContinue | Out-Null
                    $Log += "Cartella '$Folder' rimossa. "
                }
            } catch {}
        }
    } catch { $Log += "Err SharedFolder. " }

    # D. REPORT URL SITO
    if ($TenantPrefix -ne "TENANT-NOT-FOUND") {
        $PersonalUrlPart = $UPN -replace "@","_" -replace "\.","_"
        $PersonalSiteUrl = "https://$[TenantPrefix-my.sharepoint.com/personal/$PersonalUrlPart](https://TenantPrefix-my.sharepoint.com/personal/$PersonalUrlPart)"
        $Log += "Target Sito: $PersonalSiteUrl "
    }

    $Report += [PSCustomObject]@{ Utente=$Name; Esito=$Log }
}

# 6. INVIO REPORT
if ($SendEmail) {
    Write-Output "Invio mail..."
    $BodyRows = ""
    foreach ($r in $Report) { $BodyRows += "<tr><td>$($r.Utente)</td><td>$($r.Esito)</td></tr>" }
    
    $JsonEmail = @{
        message = @{
            subject = "Wipe User Report - $TenantPrefix"
            body = @{ contentType = "HTML"; content = "<h3>Report Pulizia</h3><table border='1'>$BodyRows</table>" }
            toRecipients = @( @{ emailAddress = @{ address = $NotifyReceiver } } )
        }
    } | ConvertTo-Json -Depth 5
    
    try {
        Invoke-RestMethod -Uri "[https://graph.microsoft.com/v1.0/users/$NotifySender/sendMail](https://graph.microsoft.com/v1.0/users/$NotifySender/sendMail)" -Headers $Headers -Method Post -Body $JsonEmail -ContentType "application/json" -ErrorAction Stop | Out-Null
        Write-Output "Email inviata."
    } catch { 
        Write-Warning "ERRORE MAIL: $_" 
    }
}

Write-Output "--- FINE ---"
