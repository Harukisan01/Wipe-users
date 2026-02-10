# Guida all'uso dello Script di Wipe M365 (WipeUser.ps1)

Questa guida spiega come configurare l'ambiente, i permessi necessari in Microsoft Entra (Azure AD) e come eseguire lo script per pulire gli utenti (Mailbox, OneDrive, Attività).

## 1. Prerequisiti

### Modulo Microsoft Graph PowerShell
Lo script richiede il modulo PowerShell di Microsoft Graph. Se non è installato, esegui questo comando in una console PowerShell come Amministratore:

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser -Force
```

## 2. Permessi Richiesti (Entra ID)

Lo script utilizza l'autenticazione **Delegated** (eseguito con le tue credenziali di amministratore).

### Ruoli Amministrativi Necessari
L'account che esegue lo script deve avere uno dei seguenti ruoli in Entra ID per poter cancellare dati e revocare sessioni:
-   **Global Administrator** (Consigliato per evitare problemi di accesso a OneDrive/Mailbox altrui).
-   **User Administrator** (Per revocare sessioni e gestire utenti).
-   **Exchange Administrator** (Per accedere alle mailbox, anche se potrebbe non bastare senza *Full Access*).
-   **SharePoint Administrator** (Per accedere ai siti OneDrive personali).

> [!IMPORTANT]
> **Errore 403 (Access Denied) su OneDrive**: 
> Anche con il ruolo Global Admin, l'accesso ai *contenuti* del OneDrive di un altro utente potrebbe essere bloccato di default. 
> Se ricevi un errore "403 Forbidden" su `Get-MgUserDrive`, devi aggiungere il tuo account come **Amministratore della raccolta siti** (Site Collection Admin) per il OneDrive dell'utente.
> Questo si può fare dallo **SharePoint Admin Center** > **More features** > **User profiles** > **Manage User Profiles** > Cerca utente > **Manage site collection owners**.
> Oppure via PowerShell (`Set-SPOUser`).

### Scopes (Permessi API)
Al primo avvio, lo script chiederà il consenso per i seguenti permessi:
-   `GroupMember.Read.All`: Leggere membri dei gruppi.
-   `Mail.ReadWrite`: Leggere e cancellare email *di qualsiasi utente*.
-   `Files.ReadWrite.All`: Leggere e cancellare file OneDrive *di qualsiasi utente*.
-   `User.ReadWrite.All`: Modificare utenti e revocare sessioni.

## 3. Configurazione dello Script

Apri il file `WipeUser.ps1` e verifica la sezione **CONFIGURAZIONE** in alto:

```powershell
$UserGroupId = "INSERISCI-IL-GUID-DEL-GRUPPO" 
# Esempio: "33a31c3c-b300-4879-bc15-6b6aae9c7f6e"
```
Assicurati che l'ID del gruppo sia corretto e contenga gli utenti da pulire.

## 4. Esecuzione

1.  Apri PowerShell.
2.  Spostati nella cartella dello script:
    ```powershell
    cd c:\Temp\Wipe
    ```
3.  Avvia lo script:
    ```powershell
    .\WipeUser.ps1
    ```
4.  **Login**: Si aprirà una finestra browser. Fai il login con il tuo account Amministratore (es. `admin@tuotenant.onmicrosoft.com`).
5.  **Consenso**: Se richiesto, accetta i permessi (spunta "Consent on behalf of your organization" se vuoi evitare che lo chieda ad altri admin, ma qui serve solo a te).
6.  **Conferma**: Lo script mostrerà un riepilogo.
    -   Se `$DryRun = $true`, simulerà solo le operazioni.
    -   Se `$DryRun = $false`, scrivi **ESEGUI** quando richiesto per procedere con la cancellazione.

## 5. Risoluzione Problemi

| Errore | Causa Probabile | Soluzione |
| :--- | :--- | :--- |
| `Authentication needed...` | Token scaduto o non valido. | Lo script ora forza il re-login a ogni avvio. Riprova. |
| `Get-MgUserDrive : Access denied (403)` | Non hai permessi diretti sul OneDrive dell'utente. | Aggiungi il tuo account come "Site Collection Admin" al OneDrive dell'utente (vedi sezione 2). Oppure usa un'App Registration (metodo avanzato). |
| `ResourceNotFound (404)` | La mailbox o il drive non esistono. | L'utente potrebbe non avere una licenza valida o non aver mai fatto accesso a OneDrive/Outlook. |
| `Revoke-MgUserSignInSession` fallisce | Permessi insufficienti. | Assicurati di essere **User Administrator** o **Global Admin**. |
