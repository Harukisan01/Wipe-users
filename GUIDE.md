# Guida all'utilizzo dello script di pulizia utenti Microsoft 365

Questo script automatizza la pulizia dei dati utente per un gruppo specifico di Microsoft 365. Esegue operazioni su Exchange, OneDrive e Azure AD.

## Prerequisiti

Per eseguire questo script, assicurati di avere:

1.  **PowerShell 5.1** o **PowerShell 7+** installato.
2.  **Moduli PowerShell** richiesti:
    *   `Microsoft.Graph`
    *   `Microsoft.Online.SharePoint.PowerShell`
    
    Se non presenti, lo script tenterà di installare il modulo SharePoint, ma è consigliato averli già pronti:
    ```powershell
    Install-Module Microsoft.Graph -Scope CurrentUser
    Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser
    ```

3.  **Permessi Admin**:
    *   L'utente che esegue lo script deve essere **Global Admin** o avere i ruoli combinati di **SharePoint Admin** e **User Admin**.

## Configurazione

Apri il file `cleanup_user_data.ps1` e modifica la sezione **CONFIGURAZIONE** se necessario:

*   `$UserGroupId`: Inserisci l'ID del gruppo Azure AD che contiene gli utenti da ripulire.
    *   *Default*: `"33a31c3c-b300-4879-bc15-6b6aae9c7f6e"`
*   `$DryRun`: Imposta a `$true` per simulare l'esecuzione senza cancellare nulla. Imposta a `$false` per eseguire realmente le cancellazioni.
    *   *Default*: `$false`

## Cosa fa lo script

Lo script itera su ogni utente del gruppo specificato ed esegue le seguenti azioni:

1.  **Pulizia Email**: Cancella tutti i messaggi dalla casella di posta.
2.  **Posta Eliminata**: Svuota la cartella "Deleted Items".
3.  **Cartelle OneDrive Specifiche**:
    *   Elimina ricorsivamente la cartella `/shared` (condivisi con me/da me).
    *   Elimina ricorsivamente la cartella `/favorites` (preferiti).
    *   Elimina ricorsivamente la cartella `/my` (cartella personale se presente).
    *   **Svuota il Cestino** di OneDrive (Recycle Bin).
4.  **Reset Totale OneDrive**:
    *   Elimina l'intera Site Collection dell'utente (provoca un errore 404).
    *   Rimuove il sito dal Cestino di SharePoint (eliminazione definitiva).
    *   Richiede il provisioning di un nuovo OneDrive vuoto.
5.  **Pulizia Attività**: Rimuove la cronologia attività utente.
6.  **Revoca Sessioni**: Disconnette l'utente da tutte le sessioni attive.

### Purge Definitivo Globale
Alla fine del ciclo sugli utenti, lo script esegue una scansione globale del **Cestino di SharePoint** (`Get-SPODeletedSite`) e rimuove definitivamente qualsiasi sito personale (`*-my.sharepoint.com/personal/*`) rimasto, per garantire che non ci siano residui.

## Esecuzione

1.  Apri una console PowerShell come Amministratore.
2.  Naviga nella cartella dove hai salvato lo script.
3.  Esegui:
    ```powershell
    .\cleanup_user_data.ps1
    ```
4.  Segui le istruzioni a video:
    *   Verrà richiesto il login a Microsoft Graph (via browser).
    *   Lo script tenterà di connettersi automaticamente a SharePoint Online. Se fallisce, ti chiederà l'URL Admin (es. `https://tuotenant-admin.sharepoint.com`).
    *   Dovrai digitare `ESEGUI` per confermare l'avvio delle operazioni distruttive.

## Note Importanti

*   **Irreversibilità**: Le azioni di cancellazione (email, file, siti) sono definitive. Assicurati di aver impostato correttamente il gruppo target.
*   **Tempi di attesa**: La ricreazione di un OneDrive dopo la cancellazione può richiedere da 15 minuti a 24 ore lato Microsoft, anche se lo script termina prima.
----------------------






# Guida Passo-Passo per la Gestione delle Passkey Entra ID con Azure Key Vault

Questa guida ti accompagna passo dopo passo nella configurazione e nell'utilizzo degli script per creare e gestire Passkey (FIDO2) per utenti Microsoft Entra ID (precedentemente Azure AD), utilizzando Azure Key Vault per proteggere le chiavi private.

La guida è pensata per chi non ha familiarità approfondita con Azure e spiega sia il metodo automatico (consigliato) che quello manuale per la configurazione dell'infrastruttura.

---

## Prerequisiti

Prima di iniziare, assicurati di avere quanto segue sul tuo computer:

1.  **PowerShell 7 o superiore**:
    *   Questo è essenziale. Windows PowerShell 5.1 (quello preinstallato) **non funziona**.
    *   Scarica e installa l'ultima versione da qui: [https://aka.ms/powershell](https://aka.ms/powershell)
2.  **Moduli PowerShell Necessari**:
    *   Apri PowerShell 7 come amministratore ed esegui questi comandi per installare i moduli richiesti:
        ```powershell
        Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force
        Install-Module Az.Accounts -Scope CurrentUser -Force
        ```

---

## Fase 1: Preparazione dell'Infrastruttura Azure

In questa fase creeremo le risorse necessarie su Azure: un'applicazione (Service Principal) per eseguire le operazioni e una Key Vault per custodire le chiavi.

Hai due opzioni:
*   **Opzione A (Consigliata)**: Usare lo script automatico.
*   **Opzione B (Manuale)**: Creare le risorse manualmente nel Portale Azure (se vuoi capire cosa succede "dietro le quinte" o se hai restrizioni nell'esecuzione di script).

### Opzione A: Configurazione Automatica (Consigliata)

1.  Apri PowerShell 7.
2.  Posizionati nella cartella dove hai salvato gli script (`Initialize-PasskeyKeyVault.ps1`, ecc.).
3.  Esegui il comando:
    ```powershell
    .\Initialize-PasskeyKeyVault.ps1
    ```
4.  Segui le istruzioni a schermo:
    *   Ti verrà chiesto di effettuare il login (potrebbe aprirsi una finestra del browser).
    *   Lo script creerà automaticamente tutto: App Registration, Service Principal, Key Vault e permessi.
5.  **IMPORTANTE**: Alla fine, lo script mostrerà un riepilogo con:
    *   `App ID` (ID Applicazione)
    *   `Client Secret` (Segreto Client) - **COPIALO SUBITO**, non verrà più mostrato.
    *   `Key Vault Name` (Nome Key Vault)
    *   `Tenant ID` (ID Tenant)
    Annota questi valori, ti serviranno dopo.

### Opzione B: Configurazione Manuale nel Portale Azure

Se preferisci o devi fare tutto manualmente, segui questi passaggi dettagliati:

#### 1. Creare l'App Registration
1.  Vai sul [Portale Azure](https://portal.azure.com) e cerca **Microsoft Entra ID**.
2.  Nel menu a sinistra, clicca su **App registrations** > **New registration**.
3.  Inserisci un nome (es: `KeyVault-Passkey-Service`).
4.  Lascia le impostazioni predefinite e clicca su **Register**.
5.  Dalla pagina "Overview" della nuova app, annota l'**Application (client) ID** e il **Directory (tenant) ID**.

#### 2. Creare il Client Secret
1.  Nel menu dell'app, clicca su **Certificates & secrets** > **New client secret**.
2.  Inserisci una descrizione (es: `PasskeySecret`) e scegli una scadenza (es: 12 mesi).
3.  Clicca su **Add**.
4.  **Copia subito il "Value"** del segreto. Questo è il tuo **Client Secret**.

#### 3. Assegnare i Permessi API
1.  Nel menu dell'app, clicca su **API permissions** > **Add a permission**.
2.  Seleziona **Microsoft Graph** > **Application permissions**.
3.  Cerca e seleziona il permesso: `UserAuthenticationMethod.ReadWrite.All`.
4.  Clicca su **Add permissions**.
5.  Ora clicca sul pulsante **Grant admin consent for [TuoTenant]** (richiede diritti di amministratore globale o simile) e conferma con **Yes**.

#### 4. Creare la Key Vault
1.  Cerca **Key Vaults** nella barra di ricerca del portale e clicca su **Create**.
2.  Seleziona la tua Sottoscrizione e un Resource Group (o creane uno nuovo, es: `rg-passkeys`).
3.  Inserisci un **Key vault name** univoco (es: `kv-passkey-tuonome`).
4.  Nella scheda **Access configuration**, assicurati che sia selezionato **Azure role-based access control (recommended)**.
5.  Clicca su **Review + create** e poi su **Create**.

#### 5. Assegnare il Ruolo alla Key Vault
1.  Vai alla risorsa Key Vault appena creata.
2.  Nel menu a sinistra, clicca su **Access control (IAM)** > **Add** > **Add role assignment**.
3.  Cerca e seleziona il ruolo **Key Vault Crypto Officer**. Clicca su **Next**.
4.  Clicca su **Select members**.
5.  Cerca il nome dell'App che hai creato al passaggio 1 (es: `KeyVault-Passkey-Service`) e selezionala.
6.  Clicca su **Select**, poi su **Review + assign**.

Ora hai completato manualmente la configurazione che l'Opzione A fa in automatico.

---

## Fase 2: Creazione della Passkey

Questa fase richiede operazioni crittografiche complesse (generazione chiavi, firma digitale, chiamate API Graph), quindi **deve essere eseguita tramite lo script**. Non è fattibile manualmente senza strumenti avanzati.

1.  Assicurati di avere i dati dalla Fase 1:
    *   `UserUpn`: L'email dell'utente per cui creare la passkey (es: `mario.rossi@azienda.com`).
    *   `DisplayName`: Un nome per la passkey (es: "Passkey Software").
    *   `ClientId`: L'ID dell'Applicazione.
    *   `ClientSecret`: Il segreto copiato.
    *   `KeyVaultName`: Il nome della Key Vault.
    *   `TenantId`: L'ID del Tenant.

2.  Esegui lo script di registrazione:
    ```powershell
    $secret = Read-Host -AsSecureString -Prompt "Inserisci il Client Secret"
    # (Incolla il segreto quando richiesto e premi Invio)

    .\New-KeyVaultPasskey.ps1 `
        -UserUpn "mario.rossi@azienda.com" `
        -DisplayName "Passkey Software" `
        -UseKeyVault `
        -KeyVaultName "kv-passkey-tuonome" `
        -ClientId "INSERISCI_APP_ID" `
        -ClientSecret $secret `
        -TenantId "INSERISCI_TENANT_ID"
    ```

3.  Se tutto va a buon fine, vedrai un messaggio verde di conferma. Verrà creato un file JSON nella cartella corrente (es: `mario.rossi_Passkey_Software_credential.json`). Questo file contiene i riferimenti alla chiave privata custodita nella Key Vault.

---

## Fase 3: Verifica e Accesso (Login)

Ora verifichiamo che la passkey funzioni simulando un login.

**Nota**: Dopo la registrazione, aspetta circa 30-60 secondi affinché Entra ID propaghi la nuova passkey.

1.  Esegui lo script di login puntando al file JSON creato:
    ```powershell
    $secret = Read-Host -Prompt "Inserisci il Client Secret"

    .\PasskeyLogin.ps1 `
        -KeyFilePath ".\mario.rossi_Passkey_Software_credential.json" `
        -KeyVaultClientId "INSERISCI_APP_ID" `
        -KeyVaultClientSecret $secret `
        -KeyVaultTenantId "INSERISCI_TENANT_ID"
    ```

2.  Se il login ha successo, vedrai un messaggio:
    ```
    Authentication Successful!
    User: mario.rossi@azienda.com
    Method: FIDO2 Passkey
    ```

---

## Risoluzione Problemi Comuni

*   **Errore "Attestation enforcement must be disabled"**:
    *   Devi disabilitare l'attestazione per le chiavi FIDO2 nel portale Entra ID.
    *   Vai su **Authentication methods** > **Passkey (FIDO2)** > **Configure**.
    *   Crea un profilo con "Attestation enforcement" impostato su **No** e assegnalo agli utenti o gruppi coinvolti.

*   **Errore 401/403 (Forbidden)**:
    *   Verifica di aver concesso il "Consenso Amministratore" (Admin Consent) per i permessi dell'App Registration (Fase 1, passaggio 3).
    *   Verifica che l'App abbia il ruolo "Key Vault Crypto Officer" sulla Key Vault (Fase 1, passaggio 5).

*   **Errore "Key Vault not found"**:
    *   Controlla di aver scritto correttamente il nome della Key Vault e di essere loggato nella sottoscrizione corretta (puoi usare `Connect-AzAccount`).

