# Wipe User Tool - Comprehensive Guide

![Project Logo](assets/logo.png)

This repository contains tools to securely wipe user data from Microsoft 365 (Exchange, OneDrive, SharePoint).
The solution is provided in three levels of complexity and security, ranging from local interactive execution to high-security bank-grade protection using Passkeys.

## Mental Map

*   **Level 1: Local Execution** (Interactive, Standard Security) -> `Level1_WipeUser_Local.ps1`
*   **Level 2: Automation** (Non-interactive, Service Principal) -> `Level2_WipeUser_Automation.ps1`
*   **Level 3: Bank Protection** (Phishing-Resistant, Passkey + Key Vault) -> `Level3_WipeUser_Passkey.ps1`

---

## Level 1: Local Execution

**Best for:** Ad-hoc tasks, small organizations, or testing.
**Security:** Relies on the user's interactive login (MFA supported).
**How to use:** Run the script locally on your machine. You will be prompted to sign in via a browser.

[View Level 1 Documentation](Level1_Local.md)

---

## Level 2: Automation

**Best for:** Scheduled tasks, CI/CD pipelines, Azure Automation.
**Security:** Uses an Azure App Registration (Service Principal) with Client Secret.
**How to use:** Configure an App Registration in Azure AD, grant permissions, and run the script non-interactively.

[View Level 2 Documentation](Level2_Automation.md)

---

## Level 3: Bank Protection (Passkey + Key Vault)

**Best for:** High-security environments, financial institutions, privileged access management.
**Security:** Uses **FIDO2 Passkeys** stored in **Azure Key Vault** (HSM-backed). The private key never leaves the secure vault. This is resistant to phishing and credential theft.
**How to use:**
1.  Initialize the Key Vault infrastructure using `Initialize-PasskeyKeyVault.ps1`.
2.  Register a secure passkey for your admin account using `New-KeyVaultPasskey.ps1`.
3.  Run the wipe script using `Level3_WipeUser_Passkey.ps1` which authenticates using the secure passkey.

**Source:** Based on the work by [Nathan McNulty](https://github.com/nathanmcnulty/nathanmcnulty/tree/main/Entra/passkeys/keyvault).

[View Level 3 Documentation](Level3_BankProtection.md)

---

## Repository Structure

```
/
├── README.md                           # This map
├── Level1_Local.md                     # Doc for Level 1
├── Level1_WipeUser_Local.ps1           # Script for Level 1
├── Level2_Automation.md                # Doc for Level 2
├── Level2_WipeUser_Automation.ps1      # Script for Level 2
├── Level3_BankProtection.md            # Doc for Level 3
├── Level3_WipeUser_Passkey.ps1         # Script for Level 3
├── Initialize-PasskeyKeyVault.ps1      # Helper for Level 3 setup
├── New-KeyVaultPasskey.ps1             # Helper for Level 3 registration
└── PasskeyLogin.ps1                    # Helper for Level 3 authentication
```
