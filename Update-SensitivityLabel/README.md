# Update-SensitivityLabel

Bulk migrate Microsoft 365 sensitivity labels across your entire tenant. This script scans SharePoint Online and OneDrive for Business to find all files with a specific sensitivity label and replaces it with a new one.

## Use Cases

- Migrating from pilot labels to production labels
- Consolidating multiple labels into one
- Replacing deprecated labels with new ones
- Auditing which files have specific labels (discovery mode)

## Requirements

### PowerShell Modules

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
Install-Module Microsoft.Graph -Scope CurrentUser
```

### Permissions

| Permission | Purpose |
|------------|---------|
| Data Classification Content Viewer | Required for `Export-ContentExplorerData` to scan files |
| Files.ReadWrite.All | Required for Graph API to modify files |
| Sites.ReadWrite.All | Required for Graph API to access SharePoint sites |

### Licensing

The Graph API endpoint for assigning sensitivity labels requires **premium licensing**:
- Microsoft 365 E5/A5
- Microsoft 365 E5/A5 Compliance
- Microsoft 365 E5/A5 Information Protection and Governance
- Azure Information Protection P2

Discovery mode (`-DiscoveryOnly`) works without premium licensing.

## Supported File Types

- Word documents (.docx)
- Excel spreadsheets (.xlsx)
- PowerPoint presentations (.pptx)
- PDF files (.pdf)

Legacy Office formats (.doc, .xls, .ppt) are not supported by the Microsoft Graph API.

## Usage

### Step 1: Discovery (Find files and required permissions)

```powershell
.\Update-SensitivityLabel.ps1 -OldLabelName "Confidential - Pilot" -DiscoveryOnly
```

This will:
- List all files with the specified label
- Show which sites/OneDrive locations contain labeled files
- Export results to CSV for review
- **Not require premium licensing**

### Step 2: Dry Run (Test without making changes)

```powershell
.\Update-SensitivityLabel.ps1 -OldLabelName "Confidential - Pilot" -NewLabelId "guid-here" -DryRun
```

### Step 3: Execute Migration

```powershell
.\Update-SensitivityLabel.ps1 -OldLabelName "Confidential - Pilot" -NewLabelId "guid-here"
```

## Parameters

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-OldLabelName` | Yes* | Name of the sensitivity label to replace |
| `-OldLabelId` | Yes* | GUID of the old label (alternative to OldLabelName) |
| `-NewLabelId` | Yes** | GUID of the new sensitivity label to apply |
| `-DiscoveryOnly` | No | Only list files, don't connect to Graph or make changes |
| `-DryRun` | No | Report what would be changed without making changes |
| `-Workload` | No | `ODB` (OneDrive), `SPO` (SharePoint), or `Both` (default) |
| `-JustificationText` | No | Justification text for the label change |
| `-LogPath` | No | Custom path for log file |
| `-PageSize` | No | Number of items per page (default: 100) |
| `-ThrottleDelayMs` | No | Delay between API calls in ms (default: 500) |

\* Either `-OldLabelName` or `-OldLabelId` must be provided  
\** Required unless using `-DiscoveryOnly`

## Finding Label GUIDs

The script displays all available labels at startup. You can also run:

```powershell
Connect-IPPSSession
Get-Label | Format-Table Name, DisplayName, Guid -AutoSize
```

## Output

The script generates several output files:

| File | Description |
|------|-------------|
| `LabelMigration_[timestamp].log` | Detailed execution log |
| `LabelMigration_[timestamp]_files.csv` | List of all files found |
| `LabelMigration_[timestamp]_sites.csv` | List of sites/locations |
| `LabelMigration_[timestamp]_results.csv` | Migration results (success/failure per file) |

## Important Notes

### File Metadata Changes

The Graph API will update "Modified By" and "Modified Date" on files when labels are changed. This is a Microsoft limitation and cannot be avoided.

### Content Explorer Index Lag

The Content Explorer index can be 24-48 hours behind. Files recently labeled may not appear immediately, and files already migrated may still appear in scans until the index updates.

### Throttling

The script includes built-in throttling (500ms between requests by default) to avoid Microsoft API rate limits. For large migrations, consider running in batches.

## Troubleshooting

### "PaymentRequired" Error
The account needs premium licensing (E5/AIP P2) for the `assignSensitivityLabel` API.

### No Files Found
- Verify the exact label name using `Get-Label`
- Check if you have the Data Classification Content Viewer role
- Content Explorer index may be delayed

### Permission Errors
Ensure you have access to the SharePoint sites and OneDrive locations. Global Admin does not automatically grant file access.

## License

MIT License
