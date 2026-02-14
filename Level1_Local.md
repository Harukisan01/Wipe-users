# Level 1: Local Execution

This level is designed for local execution where you run the script interactively from your machine.

## How it works

1.  **Authentication**: The script uses `Connect-MgGraph` and `Connect-SPOService` to authenticate interactively. You will be prompted to sign in with your Microsoft 365 Admin credentials.
2.  **Scope**: The script operates on a specific group of users (defined by `UserGroupId` in the script).
3.  **Actions**:
    *   Clears user's mailbox (Emails, Deleted Items).
    *   Clears OneDrive specific folders (Shared, Favorites, My).
    *   Empties OneDrive Recycle Bin.
    *   **Destructive Action**: Deletes the entire OneDrive Site Collection and recreates it empty.
    *   Clears user activities and revokes sessions.

## Prerequisites

*   PowerShell 7+ recommended.
*   Modules: `Microsoft.Graph.Authentication`, `Microsoft.Online.SharePoint.PowerShell`.
*   You must be a Global Admin or have sufficient privileges (SharePoint Admin, Exchange Admin, User Admin).

## Usage

1.  Open `Level1_WipeUser_Local.ps1`.
2.  Edit the `$UserGroupId` variable with the Object ID of the group containing the users you want to wipe.
3.  Run the script:
    ```powershell
    .\Level1_WipeUser_Local.ps1
    ```
4.  Follow the prompts to sign in.
5.  Type `EXECUTE` when asked for confirmation.

## Important Notes

*   This script is **destructive**. Data deleted from OneDrive and Mailbox may be irrecoverable if purged from Recycle Bin.
*   The script includes a "Definitive Purge" step that removes deleted personal sites from the SharePoint Recycle Bin.
