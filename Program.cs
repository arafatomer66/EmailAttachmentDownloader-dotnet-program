using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;

class Program
{
    static async Task Main(string[] args)
    {
        Log("Application started.");

        while (true)
        {
            try
            {
                Log("Starting to check emails...");
                await CheckAndDownloadEmails();
            }
            catch (System.Exception ex)
            {
                Log($"An error occurred: {ex.Message}");
            }

            Log("Waiting for the next run...");
            await Task.Delay(TimeSpan.FromMinutes(30)); // Wait for 30 minutes before the next run
        }
    }

    static async Task CheckAndDownloadEmails()
    {
        try
        {
            Application outlookApp = new Application();
            NameSpace outlookNamespace = outlookApp.GetNamespace("MAPI");
            MAPIFolder inbox = outlookNamespace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);

            Log("Outlook application initialized and inbox folder accessed.");

            // Synchronize Inbox to get the latest emails
            Log("Synchronizing inbox...");
            outlookNamespace.SendAndReceive(false);

            // Wait a few seconds to allow synchronization to complete
            await Task.Delay(TimeSpan.FromSeconds(10));

            Items mailItems = inbox.Items;

            foreach (object item in mailItems)
            {
                if (item is MailItem mailItem)
                {
                    Log($"Processing email with subject: {mailItem.Subject} and sender: {mailItem.SenderEmailAddress}");

                    if (mailItem.Subject.Contains("Email") && mailItem.SenderEmailAddress == "arafatomer66@gmail.com")
                    {
                        Log($"Matched email with subject: {mailItem.Subject}");

                        if (mailItem.Attachments.Count > 0)
                        {
                            Log($"Email has {mailItem.Attachments.Count} attachments.");

                            for (int i = 1; i <= mailItem.Attachments.Count; i++)
                            {
                                Attachment attachment = mailItem.Attachments[i];
                                string filePath = Path.Combine(@"C:\Network", attachment.FileName);
                                attachment.SaveAsFile(filePath);
                                Log($"Attachment saved to {filePath}");

                                if (i >= 2)
                                {
                                    Log("Reached limit of 2 attachments.");
                                    break;
                                }
                            }
                        }
                        else
                        {
                            Log("Email has no attachments.");
                        }
                    }
                    else
                    {
                        Log("Email does not match subject or sender.");
                    }
                }
                else
                {
                    Log("Item is not an email.");
                }
            }

            Log("Finished processing emails.");
        }
        catch (System.Exception ex)
        {
            Log($"An error occurred while reading emails or saving attachment: {ex.Message}");
        }
    }

    static void Log(string message)
    {
        string logFilePath = "log.txt";
        using (StreamWriter writer = new StreamWriter(logFilePath, true))
        {
            writer.WriteLine($"{DateTime.Now}: {message}");
        }
        Console.WriteLine($"{DateTime.Now}: {message}");
    }
}
