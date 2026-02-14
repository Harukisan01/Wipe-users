# Level 3: Bank Protection (Passkey + Key Vault)

This level provides the highest level of security by using **FIDO2 Passkeys** (phishing-resistant) where the private key is stored securely in **Azure Key Vault**.

## Architecture

*   **Azure Key Vault**: Stores the private key of the passkey.
*   **Azure App Registration**: Acts as the interface to access Key Vault (via Service Principal).
*   **Entra ID (Azure AD)**: Authenticates the user using the passkey.

## Setup Instructions

### 1. Initialize Infrastructure

Run the initialization script to set up the Key Vault, App Registration, and Permissions.

```powershell
.\Initialize-PasskeyKeyVault.ps1 -KeyVaultSku Standard -Location "eastus"
```
*   Save the **ClientId**, **ClientSecret**, and **KeyVaultName** provided in the output.

### 2. Register a Passkey

Register a new passkey for your Admin account. You will need your Admin UPN.

```powershell
$ClientSecret = Read-Host -AsSecureString -Prompt "Enter Client Secret"
.\New-KeyVaultPasskey.ps1 `
    -UserUpn "admin@yourdomain.com" `
    -DisplayName "Admin Secure Passkey" `
    -UseKeyVault `
    -KeyVaultName "kv-passkey-XXXX" `
    -ClientId "YOUR_CLIENT_ID" `
    -ClientSecret $ClientSecret `
    -TenantId "YOUR_TENANT_ID"
```
*   This will generate a `.json` credential file. **Keep this file safe** (though it doesn't contain the private key if Key Vault is used, it's still sensitive).

## Usage (Wiping User)

Once set up, you can run the Wipe User script using the Passkey authentication.

1.  Open `Level3_WipeUser_Passkey.ps1`.
2.  Set `$UserGroupId`.
3.  Run the script:

```powershell
.\Level3_WipeUser_Passkey.ps1 `
    -KeyFilePath ".\admin_passkey_credential.json" `
    -ClientId "YOUR_CLIENT_ID" `
    -ClientSecret "YOUR_CLIENT_SECRET" `
    -TenantId "YOUR_TENANT_ID"
```

The script will:
1.  Authenticate using `PasskeyLogin.ps1`.
2.  Obtain an Access Token for Microsoft Graph.
3.  Perform the wipe operations.

## Security Note

This method ensures that even if the script execution environment is compromised, the attacker cannot steal the private key because it never leaves the Azure Key Vault HSM (Hardware Security Module).
