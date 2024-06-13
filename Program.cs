using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Identity.Client;

class Program
{
    private static ExchangeService service;

    static async System.Threading.Tasks.Task Main(string[] args)
    {
        var tenantId = "256d654d-057a-4595-aaed-df8302671760"; // Your Tenant ID
        var clientId = "ef533d70-d994-47f6-a9ef-10129ea0f198"; // Your Client ID
        var clientSecret = "VB08Q~Z_ekTe9-Ha1tP1mcLtwrS_mZQQaRBy_ddf"; // Your Client Secret (Value)

        var app = ConfidentialClientApplicationBuilder.Create(clientId)
            .WithClientSecret(clientSecret)
            .WithAuthority(new Uri($"https://login.microsoftonline.com/{tenantId}"))
            .Build();

        string[] scopes = { "https://outlook.office365.com/.default" };

        AuthenticationResult result = null;
        try
        {
            result = await app.AcquireTokenForClient(scopes)
                .ExecuteAsync();
        }
        catch (MsalException ex)
        {
            Console.WriteLine($"Error acquiring access token: {ex.Message}");
            return;
        }

        service = new ExchangeService(ExchangeVersion.Exchange2013)
        {
            Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx")
        };
        service.Credentials = new OAuthCredentials(result.AccessToken);

        while (true)
        {
            await ReadEmailsAndDownloadAttachments();
            await System.Threading.Tasks.Task.Delay(TimeSpan.FromMinutes(1));
        }
    }

    private static async System.Threading.Tasks.Task ReadEmailsAndDownloadAttachments()
    {
        var inbox = Folder.Bind(service, WellKnownFolderName.Inbox);
        var view = new ItemView(10)
        {
            PropertySet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject, ItemSchema.DateTimeReceived)
        };

        SearchFilter searchFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And, new SearchFilter[]
        {
            new SearchFilter.ContainsSubstring(ItemSchema.Subject, "Start Daily"),
            new SearchFilter.ContainsSubstring(EmailMessageSchema.From, "microsoft.start@email2.microsoft.com")
        });

        var findResults = service.FindItems(inbox.Id, searchFilter, view);

        foreach (var item in findResults.OfType<EmailMessage>())
        {
            item.Load();
            if (item.HasAttachments)
            {
                int attachmentCount = 0;
                foreach (var attachment in item.Attachments.OfType<FileAttachment>())
                {
                    if (attachmentCount >= 2) break;
                    attachment.Load();
                    var filePath = Path.Combine(@"\\YourNetworkPath", attachment.Name);
                    File.WriteAllBytes(filePath, attachment.Content);
                    attachmentCount++;
                }
            }
        }
    }
}
