using Azure.Identity;
using Microsoft.Graph;

var scopes = new[] { "https://graph.microsoft.com/.default" };

var tenantId = "YOUR_TENANT_ID";
var clientId = "YOUR_APPLICATION_CLIENT_ID";
var clientSecret = "YOUR_GENERATED_SECRET";
var fromEmailAddress = "YOUR_EMAIL_ADDRESS_FROM_EXCHANGE_ONLINE";
var toEmailAddress = "YOUR_EMAIL_ADDRESS";
var saveToSentItems = true;

var options = new TokenCredentialOptions
{
	AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
};

var clientSecretCredential = new ClientSecretCredential(tenantId, clientId, clientSecret, options);

var graphClient = new GraphServiceClient(clientSecretCredential, scopes);

var message = new Message
{
	From = new Recipient
	{
		EmailAddress = new EmailAddress
		{
			Address = fromEmailAddress
		}
	},
	Subject = "Hello World from Microsoft Graph",
	Body = new ItemBody
	{
		ContentType = BodyType.Html,
		Content = "<p>Hello World from Microsoft Graph API and <b>Exchange Online</b>.</p>"
	},
	ToRecipients = new List<Recipient>()
	{
		new Recipient
		{
			EmailAddress = new EmailAddress
			{
				Address = toEmailAddress
			}
		}
	},
	CcRecipients = new List<Recipient>()
	{

	}
};

await graphClient.Users[fromEmailAddress].SendMail(message, saveToSentItems).Request().PostAsync();

Console.WriteLine("Test email sent successfully.");
Console.WriteLine("");

var messages = await graphClient.Users[fromEmailAddress].Messages.Request().Select("sender,subject,body").GetAsync();

Console.WriteLine($"Messages successfully retrieved from ({fromEmailAddress})");
Console.WriteLine("");
Console.ReadLine();