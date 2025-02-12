# Define export path
$ExportPath = "C:\ExchangeBackup"
$LogFile = "$ExportPath\ExportLog.log"
$SizeThresholdGB = 50 # Mailbox size warning threshold

# Create directories if they don't exist
if (!(Test-Path $ExportPath)) {
    New-Item -ItemType Directory -Path $ExportPath | Out-Null
}

# Function to log messages
Function Write-Log {
    param (
        [string]$Message
    )
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$Timestamp - $Message" | Out-File -Append -FilePath $LogFile
}

Write-Log "Starting mailbox export process..."

# Get all mailboxes
$Mailboxes = Get-Mailbox -ResultSize Unlimited
foreach ($Mailbox in $Mailboxes) {
    $UserExportPath = "$ExportPath\$($Mailbox.Alias)"
    
    # Create user-specific export folder
    if (!(Test-Path $UserExportPath)) {
        New-Item -ItemType Directory -Path $UserExportPath | Out-Null
    }

    # Check mailbox size
    $MailboxStats = Get-MailboxStatistics -Identity $Mailbox.Alias
    $MailboxSize = ($MailboxStats.TotalItemSize -replace '[^\d]', '') / 1GB

    if ($MailboxSize -gt $SizeThresholdGB) {
        Write-Log "WARNING: Mailbox $($Mailbox.Alias) exceeds $SizeThresholdGB GB ($MailboxSize GB)."
    }

    # Define export file paths
    $PSTFile = "$UserExportPath\$($Mailbox.Alias).pst"
    $CalendarFile = "$UserExportPath\$($Mailbox.Alias)-calendar.ics"

    # Export mailbox
    Try {
        New-MailboxExportRequest -Mailbox $Mailbox.Alias -FilePath "\\Server\Share\$($Mailbox.Alias).pst"
        Write-Log "Export initiated for mailbox: $($Mailbox.Alias)"
    } Catch {
        Write-Log "ERROR exporting mailbox $($Mailbox.Alias): $_"
    }

    # Export Address Lists
    Try {
        Get-AddressList | Export-Csv "$UserExportPath\AddressLists.csv" -NoTypeInformation
        Get-DistributionGroup | Export-Csv "$UserExportPath\DistributionGroups.csv" -NoTypeInformation
        Get-GlobalAddressList | Export-Csv "$UserExportPath\GlobalAddressList.csv" -NoTypeInformation
        Write-Log "Exported address lists for $($Mailbox.Alias)"
    } Catch {
        Write-Log "ERROR exporting address lists: $_"
    }

    # Export Calendar Data
    Try {
        New-MailboxExportRequest -Mailbox $Mailbox.Alias -FilePath "\\Server\Share\$($Mailbox.Alias)-calendar.ics"
        Write-Log "Exported calendar for $($Mailbox.Alias)"
    } Catch {
        Write-Log "ERROR exporting calendar for $($Mailbox.Alias): $_"
    }
}

Write-Log "Mailbox export process completed."
Write-Output "Mailbox export completed successfully. Check log file at $LogFile."
