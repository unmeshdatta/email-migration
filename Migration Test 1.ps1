# Define export path
$ExportPath = "K:\ExchangeBackup"
$LogFile = "$ExportPath\ExportLog.log"

# Email notification settings
$SMTPServer = "smtp.yourdomain.com"
$SMTPFrom = "admin@yourdomain.com"
$SMTPTo = "markus@yourcompany.com"
$SubjectSuccess = "Exchange Migration Completed Successfully"
$SubjectFailure = "Exchange Migration Encountered Errors"

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
$ErrorsOccurred = $false

foreach ($Mailbox in $Mailboxes) {
    $UserExportPath = "$ExportPath\$($Mailbox.Alias)"
    
    # Create user-specific export folder
    if (!(Test-Path $UserExportPath)) {
        New-Item -ItemType Directory -Path $UserExportPath | Out-Null
    }

    # Define export file paths
    $PSTFile = "$UserExportPath\$($Mailbox.Alias).pst"
    $CalendarFile = "$UserExportPath\$($Mailbox.Alias)-calendar.ics"

    # Export mailbox with retry mechanism
    $RetryCount = 0
    $MaxRetries = 3
    $ExportSuccess = $false

    while ($RetryCount -lt $MaxRetries -and -not $ExportSuccess) {
        Try {
            New-MailboxExportRequest -Mailbox $Mailbox.Alias -FilePath "$PSTFile"
            Write-Log "Export initiated for mailbox: $($Mailbox.Alias)"
            $ExportSuccess = $true
        } Catch {
            Write-Log "ERROR exporting mailbox $($Mailbox.Alias) (Attempt $($RetryCount + 1)): $_"
            $RetryCount++
            $ErrorsOccurred = $true
            Start-Sleep -Seconds 10
        }
    }

    # Export Address Lists
    Try {
        Get-AddressList | Export-Csv "$UserExportPath\AddressLists.csv" -NoTypeInformation
        Get-DistributionGroup | Export-Csv "$UserExportPath\DistributionGroups.csv" -NoTypeInformation
        Get-GlobalAddressList | Export-Csv "$UserExportPath\GlobalAddressList.csv" -NoTypeInformation
        Write-Log "Exported address lists for $($Mailbox.Alias)"
    } Catch {
        Write-Log "ERROR exporting address lists: $_"
        $ErrorsOccurred = $true
    }

    # Export Calendar Data (Netcup supports .ics files)
    Try {
        New-MailboxExportRequest -Mailbox $Mailbox.Alias -FilePath "$CalendarFile"
        Write-Log "Exported calendar for $($Mailbox.Alias)"
    } Catch {
        Write-Log "ERROR exporting calendar for $($Mailbox.Alias): $_"
        $ErrorsOccurred = $true
    }
}

Write-Log "Mailbox export process completed."

# Function to send email notifications
Function Send-EmailNotification {
    param (
        [string]$Subject,
        [string]$Body
    )
    $EmailMessage = @{
        To = $SMTPTo
        From = $SMTPFrom
        Subject = $Subject
        Body = $Body
        SmtpServer = $SMTPServer
    }
    Send-MailMessage @EmailMessage
}

# Send completion email
if ($ErrorsOccurred) {
    Write-Log "Migration completed with errors. Sending failure notification."
    Send-EmailNotification -Subject $SubjectFailure -Body "Exchange migration completed with errors. Please check the log file at $LogFile for details."
} else {
    Write-Log "Migration completed successfully. Sending success notification."
    Send-EmailNotification -Subject $SubjectSuccess -Body "Exchange migration completed successfully. All mailboxes were exported without issues."
}

Write-Output "Mailbox export completed successfully. Check log file at $LogFile."

