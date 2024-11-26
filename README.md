# Exchange Online Mailbox Item Counter

This PowerShell script is an essential tool for Exchange Online mailbox migrations and capacity planning. It provides detailed analysis of mailbox contents, helping administrators:

- Compare source and target mailbox sizes and item counts during migrations
- Verify successful data transfer by comparing pre and post-migration statistics
- Plan storage requirements for mailbox migrations
- Identify large mailboxes that may require special handling during migration
- Track migration progress by comparing item counts
- Generate comprehensive reports for migration documentation

The script retrieves and analyzes item counts and sizes across different folders in Exchange Online mailboxes, making it invaluable for both individual and bulk mailbox migrations.

## Prerequisites

- Exchange Online PowerShell Module (`ExchangeOnlineManagement`)
- Exchange Online Admin credentials
- PowerShell 5.1 or later

## Installation

1. Install the Exchange Online PowerShell module if not already installed:
```powershell
Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber
```

2. Download the script to your preferred location
3. Ensure you have the necessary permissions (Exchange Administrator role recommended)

## Features

- **Detailed Mailbox Analysis**
  - Individual mailbox statistics with folder-by-folder breakdown
  - Bulk processing capability for organization-wide analysis
  - Automated size formatting (B, KB, MB, GB)
  - CSV export functionality for data analysis
  - Filtered folder view focusing on essential mailbox components
  - Total size and item count calculations

## Folder Coverage

The script analyzes the following folders:
- Inbox
- Sent Items
- Deleted Items
- Archive
- Drafts
- Junk Email
- Outbox
- Calendar (main folder only)
- Contacts (main folder only)

## Usage

### Option 1: Individual Mailbox Check

```powershell
.\Get-MailboxItemCounts.ps1
# Select option 1 when prompted
# Enter the email address of the mailbox to analyze
```

Example output:
```
Mailbox Statistics for: user@domain.com

Folder                 Items    Size
------                 -----    ----
/Inbox                 1250     2.5 GB
/Sent Items           850      1.8 GB
/Deleted Items        320      750 MB
/Archive              2500     5.2 GB
...

Total Folders: 9
Total Items: 5420
```

### Option 2: Bulk Processing

```powershell
.\Get-MailboxItemCounts.ps1
# Select option 2 when prompted
```

The script will:
1. Connect to Exchange Online
2. Process all mailboxes in the organization
3. Generate a CSV report with the filename format: `MailboxStats_YYYY-MM-DD.csv`

## CSV Report Format

The generated CSV includes the following columns:
- UserPrincipalName
- DisplayName
- TotalItems
- TotalSizeGB
- InboxItems
- SentItemsCount
- DeletedItemsCount
- ArchiveItems
- LastLogonTime

## Error Handling

The script includes robust error handling:
- Validates Exchange Online connection
- Checks for required permissions
- Handles mailbox access errors gracefully
- Provides detailed error messages for troubleshooting

## Best Practices

1. Run during off-peak hours for bulk processing
2. Use a privileged admin account with appropriate permissions
3. Monitor script execution for large organizations
4. Review CSV output for any anomalies
5. Keep the Exchange Online PowerShell module updated

## Limitations

- Processing time increases with mailbox size and count
- Rate limiting may apply for large organizations
- Some folder types are intentionally excluded (Tasks, Notes, Journal, SyncIssues)
- Only processes main Calendar and Contacts folders (not subfolders)

## Troubleshooting

Common issues and solutions:

1. **Connection Errors**
   - Verify internet connectivity
   - Check admin credentials
   - Ensure MFA is properly configured

2. **Permission Issues**
   - Verify Exchange Administrator role
   - Check for conditional access policies
   - Ensure no custom RBAC restrictions

3. **Performance Issues**
   - Run during off-peak hours
   - Process smaller batches
   - Check network latency

## Contributing

Feel free to submit issues and enhancement requests!

## License

This script is released under the MIT License.

## Author

Created by: Mezba Uddin
Last Updated: 2024

## Output

The script provides:
- Folder-by-folder breakdown
- Item counts per folder
- Size of each folder (automatically formatted to appropriate unit)
- Total number of folders
- Total item count
- Total size across all folders

## Notes

- The script automatically connects to Exchange Online
- Admin credentials are required for execution
- Excludes task folders, notes, journal items, and sync issues
- Size formatting is automatic (B, KB, MB, GB)
