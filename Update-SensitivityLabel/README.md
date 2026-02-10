# Update-SensitivityLabel

Bulk migrate Microsoft 365 sensitivity labels across SharePoint Online and OneDrive for Business. Scans the entire tenant using Purview Content Explorer, then uses the Microsoft Graph API to replace labels programmatically.

## Scripts

| Script | Description |
|--------|-------------|
| `Update-SensitivityLabel.ps1` | Main migration script. Discovers files with a specific sensitivity label and replaces it with a new one. |
| `Manage-SPOAdmin.ps1` | Optional helper script to bulk add/remove Site Collection Admin permissions on SharePoint sites. Only needed if using interactive login instead of an app registration. |

## Prerequisites

### Modules

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
Install-Module Microsoft.Graph -Scope CurrentUser
Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser  # For Manage-SPOAdmin.ps1
```

### Roles & Permissions

- **Data Classification Content Viewer** role (for `Export-ContentExplorerData`)
- An **App Registration** (Confidential Client) with:
  - **Application permissions** (not Delegated): `Files.ReadWrite.All`, `Sites.ReadWrite.All`
  - **Admin consent** granted

### Metered API (Required)

The `assignSensitivityLabel` Graph API endpoint is a [metered/premium API](https://learn.microsoft.com/en-us/graph/metered-api-setup). You must:

1. Have an active **Azure subscription**
2. Create a **Microsoft.GraphServices/accounts** resource linking your app registration to the subscription
3. Use a **Confidential Client** (app + client secret) -- interactive login will return `402 Payment Required`

```bash
# Register the provider (one-time)
az provider register --namespace Microsoft.GraphServices

# Create the billing resource
az graph-services account create \
  --resource-group <resource-group> \
  --resource-name GraphServicesMeteredAPI \
  --subscription <subscription-id> \
  --app-id <app-client-id> \
  --location global
```

**Cost:** ~$0.00185 per API call (~$0.37 for 200 files).

## Usage

### With App Registration (recommended)

An app registration with Application permissions has tenant-wide access to all SharePoint sites and OneDrive locations, so no additional permission grants are needed.

#### Step 1: Discovery (find files with the old label)

No Graph connection or app registration needed for this step -- uses your interactive Purview session.

```powershell
.\Update-SensitivityLabel.ps1 -OldLabelName "Confidential - Old" -DiscoveryOnly
```

Outputs CSV files listing all files and sites with the label.

#### Step 2: Dry run (test without changes)

```powershell
.\Update-SensitivityLabel.ps1 `
  -OldLabelName "Confidential - Old" `
  -NewLabelId "new-label-guid-here" `
  -TenantId "<tenant-id>" `
  -ClientId "<app-client-id>" `
  -ClientSecret "<client-secret>" `
  -DryRun
```

#### Step 3: Live migration

```powershell
.\Update-SensitivityLabel.ps1 `
  -OldLabelName "Confidential - Old" `
  -NewLabelId "new-label-guid-here" `
  -TenantId "<tenant-id>" `
  -ClientId "<app-client-id>" `
  -ClientSecret "<client-secret>"
```

### With Interactive Login (legacy)

If you're not using an app registration, you'll need to grant your user account Site Collection Admin access to each site containing files. Use `Manage-SPOAdmin.ps1` for this:

```powershell
# Before migration: grant access
.\Manage-SPOAdmin.ps1 `
  -AdminUrl "https://contoso-admin.sharepoint.com" `
  -UserEmail "admin@contoso.com" `
  -CsvPath ".\sites.csv" `
  -Action Add

# Run migration (interactive login -- will get 402 on label changes without metered API)
.\Update-SensitivityLabel.ps1 `
  -OldLabelName "Confidential - Old" `
  -NewLabelId "new-label-guid-here"

# After migration: revoke access
.\Manage-SPOAdmin.ps1 `
  -AdminUrl "https://contoso-admin.sharepoint.com" `
  -UserEmail "admin@contoso.com" `
  -CsvPath ".\sites.csv" `
  -Action Remove
```

> **Note:** Interactive login does not work with the metered API. Label changes will fail with `402 Payment Required`. Use an app registration instead.

## Parameters (Update-SensitivityLabel.ps1)

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-OldLabelName` | Yes* | Name of the sensitivity label to find and replace |
| `-OldLabelId` | Yes* | GUID of the old label (alternative to Name) |
| `-NewLabelId` | Yes** | GUID of the new sensitivity label to apply |
| `-TenantId` | Yes*** | Azure AD tenant ID |
| `-ClientId` | Yes*** | App registration client ID |
| `-ClientSecret` | Yes*** | App registration client secret |
| `-DryRun` | No | Report what would change without making changes |
| `-DiscoveryOnly` | No | Only list files and sites, skip Graph entirely |
| `-Workload` | No | `SPO`, `ODB`, or `Both` (default: `Both`) |
| `-JustificationText` | No | Audit justification text for label changes |
| `-PageSize` | No | Content Explorer page size (default: 100) |
| `-ThrottleDelayMs` | No | Delay between API calls in ms (default: 500) |
| `-LogPath` | No | Custom log file path |

\* One of `-OldLabelName` or `-OldLabelId` is required.  
\*\* Required unless using `-DiscoveryOnly`.  
\*\*\* Required for live label changes (metered API requires confidential client).

## Parameters (Manage-SPOAdmin.ps1)

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-AdminUrl` | Yes | SharePoint admin center URL |
| `-UserEmail` | Yes | User/account to add or remove as admin |
| `-CsvPath` | Yes | Path to CSV file with site URLs |
| `-Action` | Yes | `Add` or `Remove` |

## Supported File Types

The Graph API `assignSensitivityLabel` endpoint supports:
- `.docx` (Word)
- `.xlsx` (Excel)
- `.pptx` (PowerPoint)
- `.pdf` (PDF)

Other file types (e.g., `.page` Loop files) are not supported and will be skipped.

## Known Limitations

- **Metered API is required** -- without it, all label changes return `402 Payment Required`
- **Only confidential clients** (app + secret) work with metered APIs -- interactive login does not
- **Azure managed identities** are not supported for metered APIs
- **File metadata changes** -- Graph API updates "Modified By" and "Modified Date" when changing labels
- **Async processing** -- `assignSensitivityLabel` is asynchronous; files may be temporarily locked during processing
- **Loop pages** (`.page`) are not supported by the API
- **Only available in Microsoft global environment** -- not available in GCC/national clouds

## Outputs

The script generates the following files in the script directory:

| File | Mode | Description |
|------|------|-------------|
| `LabelMigration_<timestamp>.log` | All | Detailed execution log |
| `LabelMigration_<timestamp>_files.csv` | Discovery | All files found with the label |
| `LabelMigration_<timestamp>_sites.csv` | Discovery | Unique sites/locations |
| `LabelMigration_<timestamp>_results.csv` | Migration | Per-file results with status |

## Troubleshooting

| Error | Cause | Fix |
|-------|-------|-----|
| `402 Payment Required` | Metered API not enabled | Create `Microsoft.GraphServices/accounts` resource in Azure |
| `403 Forbidden` | Missing permissions or no admin consent | Check app registration has Application permissions with admin consent |
| `401 Unauthorized` | Bad credentials or expired secret | Verify TenantId/ClientId/ClientSecret values |
| `400 Bad Request` | Invalid label ID or unsupported file type | Verify the NewLabelId GUID is correct; check file type is supported |

## License

MIT
