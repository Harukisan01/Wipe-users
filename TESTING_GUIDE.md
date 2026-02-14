# Testing Guide

This guide explains how to safely test the Wipe User Tool scripts without risking data loss.

## Disclaimer

**These scripts are destructive by design.** They are intended to permanently delete user data (emails, OneDrive files, SharePoint sites).
While we have implemented a "Dry Run" mode, **always test in a non-production environment** (e.g., a dev tenant) or with a dedicated test user first.

## Dry Run Mode

All three levels (Level 1, Level 2, Level 3) include a `$DryRun` variable at the top of the script.

*   `$DryRun = $true`: The script will **simulate** the actions (logging what would happen) but will **NOT** execute any deletion commands.
*   `$DryRun = $false` (Default): The script **WILL** execute deletions.

### How to use Dry Run

1.  Open the script you want to test (`Level1_WipeUser_Local.ps1`, etc.).
2.  Locate the line `$DryRun = $false`.
3.  Change it to `$DryRun = $true`.
4.  Run the script.
5.  Review the output console. You should see messages prefixed with `[DRY-RUN]`.

## Level-Specific Testing Instructions

### Level 1: Local Execution

1.  Set `$DryRun = $true`.
2.  Set `$UserGroupId` to a group containing a test user.
3.  Run `.\Level1_WipeUser_Local.ps1`.
4.  Follow the interactive login prompts.
5.  Verify that it connects to Graph and SharePoint and lists the items it *would* delete.

### Level 2: Automation

1.  **Prerequisite:** You must have an App Registration with the required permissions (see [Level2_Automation.md](Level2_Automation.md)).
2.  Set `$DryRun = $true`.
3.  Fill in `$TenantId`, `$ClientId`, `$ClientSecret` (or pass them as parameters).
4.  Run `.\Level2_WipeUser_Automation.ps1`.
5.  Verify successful non-interactive authentication and simulated wipe actions.

### Level 3: Bank Protection (Passkeys)

**Note:** This level requires significant setup.

1.  **Setup Infrastructure:**
    *   Run `.\Initialize-PasskeyKeyVault.ps1` (it defaults to `westeurope`).
    *   Verify that a Key Vault and App Registration are created in your Azure subscription.
2.  **Register a Passkey:**
    *   Run `.\New-KeyVaultPasskey.ps1` for your test admin account.
    *   This registers a credential in Entra ID and stores the private key in the Key Vault.
3.  **Test Wipe:**
    *   Set `$DryRun = $true` in `Level3_WipeUser_Passkey.ps1`.
    *   Run the script pointing to your generated credential file.
    *   The script should authenticate using the Passkey (accessing Key Vault) and then simulate the wipe.

## Verification Checklist

- [ ] **Authentication:** Does the script successfully connect to Microsoft Graph and SharePoint Online?
- [ ] **User Discovery:** Does it correctly identify the users in the target group?
- [ ] **Dry Run Safety:** When `$DryRun` is true, are deletion commands skipped?
- [ ] **Scope:** Does it target *only* the specific user's resources (e.g., their personal site) and not others?
