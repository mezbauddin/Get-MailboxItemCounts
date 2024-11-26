# Get-MailboxItemCounts.ps1

## Overview

**Get-MailboxItemCounts.ps1** is a PowerShell script designed for **Exchange Online** mailbox management. It is especially useful for **migration validation**, allowing administrators to compare item counts between source and destination tenants.

## Features

- **Analyze Individual Mailboxes**: View folder sizes and item counts interactively.
- **Export Statistics for All Mailboxes**: Generate a CSV report for all mailboxes.
- **Migration Validation**: Compare source and destination tenants to verify migration success.

## Requirements

- **Modules**:
  - Install the Exchange Online Management module:
    ```powershell
    Install-Module ExchangeOnlineManagement
    ```
- **Permissions**:
  - Admin access to Exchange Online with mailbox query permissions.

## How to Use

1. **Run the Script**:
   ```powershell
   .\Get-MailboxItemCounts.ps1
