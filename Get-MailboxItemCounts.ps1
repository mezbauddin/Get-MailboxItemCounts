# Import the Exchange Online PowerShell module
Import-Module ExchangeOnlineManagement

# Connect to Exchange Online
# You may need to provide your credentials here
Connect-ExchangeOnline #Admin account as every time this script will run

# Function to get all mailbox folders
function Get-AllMailboxFolders {
    param (
        [string]$Mailbox
    )
    try {
        # Filter for email and outlook-related folders only
        $allFolders = Get-MailboxFolderStatistics -Identity $Mailbox | 
            Where-Object { 
                (
                    $_.FolderPath -match "/Inbox|/Sent|/Deleted|/Archive|/Drafts|/Junk|/Outbox" -or
                    # Only include main Calendar and Contacts folders, not subfolders
                    $_.FolderPath -eq "/Calendar" -or
                    $_.FolderPath -eq "/Contacts"
                ) -and
                $_.FolderType -ne "Tasks" -and 
                $_.FolderType -ne "Notes" -and 
                $_.FolderType -ne "Journal" -and
                $_.FolderType -ne "SyncIssues"
            } |
            Select-Object FolderPath, ItemsInFolder, FolderSize |
            Where-Object { $_.ItemsInFolder -ge 0 }
        return $allFolders
    } catch {
        Write-Host "Error retrieving folders for mailbox: $Mailbox" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        return $null
    }
}

# Function to format size
function Format-FolderSize {
    param (
        [string]$SizeString
    )
    if ($SizeString -match "\(([\d,]+) bytes\)") {
        $bytes = [long]($matches[1] -replace ',','')
        if ($bytes -ge 1GB) {
            return "$([math]::Round($bytes/1GB, 2)) GB"
        } elseif ($bytes -ge 1MB) {
            return "$([math]::Round($bytes/1MB, 2)) MB"
        } elseif ($bytes -ge 1KB) {
            return "$([math]::Round($bytes/1KB, 2)) KB"
        } else {
            return "$bytes B"
        }
    }
    return $SizeString
}

# Function to display mailbox statistics
function Show-MailboxStats {
    param (
        [string]$Mailbox
    )
    
    Write-Host "`nMailbox Statistics for: $Mailbox`n" -ForegroundColor Green
    
    $allFolders = Get-AllMailboxFolders -Mailbox $Mailbox
    if ($null -eq $allFolders) {
        return
    }

    $results = @()
    $totalItems = 0
    $totalSizeBytes = 0

    foreach ($folder in $allFolders) {
        # Extract bytes from folder size string
        if ($folder.FolderSize -match "\(([\d,]+) bytes\)") {
            $bytes = [long]($matches[1] -replace ',','')
            $totalSizeBytes += $bytes
        }
        
        $totalItems += $folder.ItemsInFolder
        
        $results += [PSCustomObject]@{
            'Folder' = $folder.FolderPath
            'Items' = $folder.ItemsInFolder
            'Size' = Format-FolderSize -SizeString $folder.FolderSize
        }
    }

    # Add total row
    $results += [PSCustomObject]@{
        'Folder' = "TOTAL"
        'Items' = $totalItems
        'Size' = Format-FolderSize -SizeString "($totalSizeBytes bytes)"
    }

    $results | Format-Table -AutoSize
    
    Write-Host "Total Folders: $($allFolders.Count)" -ForegroundColor Yellow
    Write-Host "Total Items: $totalItems" -ForegroundColor Yellow
    Write-Host "Total Size: $(Format-FolderSize -SizeString "($totalSizeBytes bytes)")" -ForegroundColor Yellow
}

# Function to search for a mailbox by name or email
function Search-Mailbox {
    param (
        [string]$SearchTerm
    )
    $mailboxes = Get-Mailbox -ResultSize Unlimited | Where-Object {
        $_.DisplayName -like "*$SearchTerm*" -or $_.PrimarySmtpAddress -like "*$SearchTerm*"
    }
    return $mailboxes
}

# Prompt user for option
$option = Read-Host "Enter 1 to check individual mailbox, 2 for all mailboxes to CSV"

if ($option -eq "1") {
    # Option 1: Search and display individual user mailbox item count
    $searchTerm = Read-Host "Enter the name or email to search"
    $mailboxes = Search-Mailbox -SearchTerm $searchTerm
    if ($mailboxes.Count -eq 0) {
        Write-Host "No mailboxes found for the search term: $searchTerm"
    } elseif ($mailboxes.Count -eq 1) {
        Show-MailboxStats -Mailbox $mailboxes[0].PrimarySmtpAddress
    } else {
        Write-Host "Multiple mailboxes found. Please select one of the following:"
        for ($i = 0; $i -lt $mailboxes.Count; $i++) {
            Write-Host "[$i] $($mailboxes[$i].DisplayName) - $($mailboxes[$i].PrimarySmtpAddress)"
        }
        $selection = Read-Host "Enter the number corresponding to the mailbox"
        if ($selection -ge 0 -and $selection -lt $mailboxes.Count) {
            Show-MailboxStats -Mailbox $mailboxes[$selection].PrimarySmtpAddress
        } else {
            Write-Host "Invalid selection. Please run the script again."
        }
    }
} elseif ($option -eq "2") {
    # Option 2: All users mailbox item count to CSV
    $allMailboxes = Get-Mailbox -ResultSize Unlimited
    $results = @()
    $grandTotalItems = 0
    $grandTotalSizeBytes = 0
    $grandTotalFolders = 0
    
    foreach ($mbx in $allMailboxes) {
        Write-Host "Processing mailbox: $($mbx.PrimarySmtpAddress)" -ForegroundColor Yellow
        
        $allFolders = Get-AllMailboxFolders -Mailbox $mbx.PrimarySmtpAddress
        if ($null -eq $allFolders) {
            continue
        }

        $totalItems = 0
        $totalSizeBytes = 0
        $folderStats = @{}

        foreach ($folder in $allFolders) {
            if ($folder.FolderSize -match "\(([\d,]+) bytes\)") {
                $bytes = [long]($matches[1] -replace ',','')
                $totalSizeBytes += $bytes
            }
            $totalItems += $folder.ItemsInFolder
            $folderStats[$folder.FolderPath] = @{
                'Items' = $folder.ItemsInFolder
                'Size' = Format-FolderSize -SizeString $folder.FolderSize
            }
        }

        # Update grand totals
        $grandTotalItems += $totalItems
        $grandTotalSizeBytes += $totalSizeBytes
        $grandTotalFolders += $allFolders.Count

        $result = [PSCustomObject]@{
            'Mailbox' = $mbx.PrimarySmtpAddress
            'Total Folders' = $allFolders.Count
            'Total Items' = $totalItems
            'Total Size' = Format-FolderSize -SizeString "($totalSizeBytes bytes)"
        }

        # Add folder-specific information
        foreach ($folder in $allFolders) {
            $folderPath = $folder.FolderPath
            $result | Add-Member -NotePropertyName "${folderPath} Items" -NotePropertyValue $folderStats[$folderPath]['Items']
            $result | Add-Member -NotePropertyName "${folderPath} Size" -NotePropertyValue $folderStats[$folderPath]['Size']
        }

        $results += $result
    }

    # Add grand total row
    $grandTotal = [PSCustomObject]@{
        'Mailbox' = "GRAND TOTAL (All Users)"
        'Total Folders' = $grandTotalFolders
        'Total Items' = $grandTotalItems
        'Total Size' = Format-FolderSize -SizeString "($grandTotalSizeBytes bytes)"
    }

    # Add empty values for folder-specific columns in grand total row
    $firstResult = $results[0]
    $folderColumns = $firstResult.PSObject.Properties.Name | Where-Object { 
        $_ -notlike "Total*" -and $_ -ne "Mailbox" 
    }
    foreach ($column in $folderColumns) {
        $grandTotal | Add-Member -NotePropertyName $column -NotePropertyValue "" -Force
    }

    $results += $grandTotal

    $results | Export-Csv -Path "MailboxItemCounts.csv" -NoTypeInformation
    Write-Host "`nItem counts saved to MailboxItemCounts.csv" -ForegroundColor Green
    Write-Host "Total mailboxes processed: $($results.Count - 1)" -ForegroundColor Yellow
    Write-Host "`nGrand Totals:" -ForegroundColor Cyan
    Write-Host "Total Folders across all mailboxes: $grandTotalFolders" -ForegroundColor Cyan
    Write-Host "Total Items across all mailboxes: $grandTotalItems" -ForegroundColor Cyan
    Write-Host "Total Size across all mailboxes: $(Format-FolderSize -SizeString "($grandTotalSizeBytes bytes)")" -ForegroundColor Cyan
} else {
    Write-Output "Invalid option selected. Please enter 1 or 2."
}

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false
