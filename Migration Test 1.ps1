# Define export path
$ExportPath = "C:\ExchangeBackup"

# Create directory if it doesn't exist
if (!(Test-Path $ExportPath)) { 
    New-Item -ItemType Directory -Path $ExportPath 
}

# Export all user mailboxes to PST
$Mailboxes = Get-Mailbox -ResultSize Unlimited 
foreach ($Mailbox in $Mailboxes) { 
    $PSTFile = "$ExportPath\$($Mailbox.Alias).pst" 
    New-MailboxExportRequest -Mailbox $Mailbox.Alias -FilePath "\\Server\Share\$($Mailbox.Alias).pst" 
}

# Export Address Lists
Get-AddressList | Export-Csv "$ExportPath\AddressLists.csv" -NoTypeInformation 
Get-DistributionGroup | Export-Csv "$ExportPath\DistributionGroups.csv" -NoTypeInformation 
Get-GlobalAddressList | Export-Csv "$ExportPath\GlobalAddressList.csv" -NoTypeInformation

# Export Calendar Data
foreach ($Mailbox in $Mailboxes) { 
    $ICSFile = "$ExportPath\$($Mailbox.Alias).ics" 
    New-MailboxExportRequest -Mailbox $Mailbox.Alias -FilePath "\\Server\Share\$($Mailbox.Alias)-calendar.ics" 
}

Write-Output "Mailbox export completed successfully."
 