# Function to log messages to a file and console
function Log($message) {
    $logFilePath = "log.txt"
    Add-Content -Path $logFilePath -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'): $message"
    Write-Output "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'): $message"
}

# Main loop to check emails every 30 minutes
while ($true) {
    try {
        Log "Starting to check emails..."
        
        # Initialize Outlook application
        $outlookApp = New-Object -ComObject Outlook.Application
        $namespace = $outlookApp.GetNamespace("MAPI")
        $inbox = $namespace.GetDefaultFolder(6)  # 6 represents the Inbox folder
        
        Log "Outlook application initialized and inbox folder accessed."

        # Synchronize Inbox to get the latest emails
        Log "Synchronizing inbox..."
        $namespace.SendAndReceive($false)

        # Wait a few seconds to allow synchronization to complete
        Start-Sleep -Seconds 10

        $mailItems = $inbox.Items
        
        foreach ($item in $mailItems) {
            if ($item -is [Microsoft.Office.Interop.Outlook.MailItem]) {
                $mailItem = $item

                Log "Processing email with subject: $($mailItem.Subject) and sender: $($mailItem.SenderEmailAddress)"

                if ($mailItem.Subject -like "Email" -and $mailItem.SenderEmailAddress -eq "arafatomer66@gmail.com") {
                    Log "Matched email with subject: $($mailItem.Subject)"

                    if ($mailItem.Attachments.Count -gt 0) {
                        Log "Email has $($mailItem.Attachments.Count) attachments."

                        $attachmentsProcessed = 0
                        foreach ($attachment in $mailItem.Attachments) {
                            $filePath = "C:\Network\" + $attachment.FileName
                            $attachment.SaveAsFile($filePath)
                            Log "Attachment saved to $filePath"

                            $attachmentsProcessed++
                            if ($attachmentsProcessed -ge 2) {
                                Log "Reached limit of 2 attachments."
                                break
                            }
                        }
                    } else {
                        Log "Email has no attachments."
                    }
                } else {
                    Log "Email does not match subject or sender."
                }
            } else {
                Log "Item is not an email."
            }
        }

        Log "Finished processing emails."

    } catch {
        Log "An error occurred: $($_.Exception.Message)"
    }

    Log "Waiting for the next run..."
    Start-Sleep -Seconds (30 * 60)  # Wait for 30 minutes before the next run
}
