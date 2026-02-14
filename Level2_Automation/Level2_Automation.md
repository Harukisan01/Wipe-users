# Level 2: Azure Automation Account

This level is designed for automated execution using an Azure Automation Account and App Registration (Service Principal).

## How it works

1.  **Authentication**: The script uses an App Registration (Service Principal) to authenticate non-interactively.
2.  **Environment**: It is designed to run in an Azure Automation Runbook or a CI/CD pipeline.
3.  **Permissions**: The App Registration requires API permissions (Application Permissions) for Graph and SharePoint.

## Prerequisites

1.  **App Registration**:
    *   Create an App Registration in Microsoft Entra ID.
    *   Create a Client Secret.
    *   Grant the following **Application Permissions**:
        *   **Microsoft Graph**: `User.ReadWrite.All`, `Files.ReadWrite.All`, `Sites.ReadWrite.All`, `Mail.ReadWrite`, `GroupMember.Read.All`, `Application.Read.All`.
        *   **SharePoint**: `Sites.FullControl.All` (or assign `SharePoint Administrator` role to the Service Principal).
    *   **Grant Admin Consent** for the permissions.

2.  **Automation Account (Optional)**:
    *   Create an Azure Automation Account.
    *   Import required modules: `Microsoft.Graph.Authentication`, `Microsoft.Online.SharePoint.PowerShell`.
    *   Create Variables or Credentials for `TenantId`, `ClientId`, and `ClientSecret`.

## Usage

1.  Open `Level2_WipeUser_Automation.ps1`.
2.  Set the following variables at the top of the script (or pass them if modifying for parameters):
    *   `$TenantId`: Your Tenant ID.
    *   `$ClientId`: Your App Registration Application (Client) ID.
    *   `$ClientSecret`: Your App Registration Client Secret.
    *   `$UserGroupId`: The Object ID of the group to process.
3.  Run the script.

## Important Notes

*   This script runs with **Application Permissions**, meaning it has broad access to the tenant. Ensure the App Registration credentials are secured.
*   For SharePoint operations, it uses `Connect-SPOService` with an Access Token. Ensure the `Microsoft.Online.SharePoint.PowerShell` module is up to date.
