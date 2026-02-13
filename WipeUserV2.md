# 📘 Guida Completa: Automazione Wipe User su Azure (v7.2)

Questa guida descrive come configurare un'automazione su Microsoft Azure per eliminare dati utente (Email, OneDrive, Siti Personali) in modo massivo partendo da un Gruppo di Sicurezza.

**Metodo:** API REST Diretta (Nessun modulo PowerShell esterno richiesto).
**Resilienza:** Anti-blocco, esecuzione non interattiva.

---

## 🚀 Fase 1: Registrazione App su Entra ID

Per permettere allo script di agire sul tenant senza login manuale, dobbiamo creare un'identità digitale (Service Principal).

1. Accedi al [Portale Azure](https://portal.azure.com).
2. Vai su **Microsoft Entra ID**.
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
WipeUserV2.ps1
