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
