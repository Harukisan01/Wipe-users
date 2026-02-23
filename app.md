# Permessi necessari per App Registration

## Script WIPE-USER -- Azure Automation (API REST pura)

## Modalità di autenticazione

Lo script utilizza: - OAuth2 client_credentials - Application
permissions (non Delegated) - Microsoft Graph API

È necessario concedere Admin Consent.

------------------------------------------------------------------------

## 1. Lettura membri del gruppo

Endpoint utilizzato: GET /groups/{id}/members

Permessi richiesti (Application):

-   Group.Read.All
-   User.Read.All

Alternativa più restrittiva: - GroupMember.Read.All

------------------------------------------------------------------------

## 2. Cancellazione email utente

Endpoint utilizzati: GET /users/{id}/messages\
DELETE /users/{id}/messages/{id}

Permesso richiesto:

-   Mail.ReadWrite.All

------------------------------------------------------------------------

## 3. Eliminazione file OneDrive

Endpoint utilizzati: GET /users/{id}/drive/root/children\
DELETE /users/{id}/drive/items/{id}

Permesso richiesto:

-   Files.ReadWrite.All

------------------------------------------------------------------------

## 4. Invio email report

Endpoint utilizzato: POST /users/{sender}/sendMail

Permesso richiesto:

-   Mail.Send

Nota: il sender deve avere una mailbox valida.

------------------------------------------------------------------------

## 5. (Opzionale) Gestione SharePoint via Graph

Permesso richiesto:

-   Sites.ReadWrite.All

------------------------------------------------------------------------

# Riepilogo completo permessi Application

-   Group.Read.All
-   User.Read.All
-   Mail.ReadWrite.All
-   Files.ReadWrite.All
-   Mail.Send
-   (Opzionale) Sites.ReadWrite.All

Tutti devono essere configurati come: Application permissions + Admin
consent

------------------------------------------------------------------------

# Requisiti Azure Automation

L'Automation Account deve contenere le seguenti variabili:

-   TenantId
-   AppClientId
-   AppClientSecret
-   TargetUserGroupId
-   (Opzionale) NotificationSender
-   (Opzionale) NotificationReceiver

Deve inoltre avere: - Accesso in uscita verso
https://graph.microsoft.com - Nessuna Conditional Access che blocchi il
client credential flow
