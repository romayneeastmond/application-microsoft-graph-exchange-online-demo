# Microsoft Graph API and Exchange Online

This project simply connects to an Exchange Online instance by using an application defined in Azure Active Directory (AD). The application is given access to the Microsoft Graph API with Mail.Read and Mail.Send permissions. These permissions allow for the console application to connect to any valid Azure AD Exchance Online mailbox.

## How to Use

1. Create a new application in Azure Active Directory. 
1. Create a secret value and copy this value immediately after its creation
1. Create API permissions for Mail.Read and Mail.Send which are found in the Microsoft Graph API
1. Grant admin consent to both permissions defined in the previous step
1. Edit the Program.cs file and replace the associated variables 

Note that the "Overview" blade, of the created application, contains the Application (client) ID and Directory (tenant) ID values. 

This demonstration uses a secret value which can be found in the "Certificates & secrets" blade.

The "API permissions" blade defines the two necessary types for Mail.Read and Mail.Send which both need to be granted admin consent. 

The fromEmailAddress must be for an active user in Exchange Online.

## Useful Links

1. [Create a Microsoft Graph Client](https://learn.microsoft.com/en-us/graph/sdks/create-client?tabs=CS) 
1. [Choose a Microsoft Graph Authentication Provider Based on Scenario](https://learn.microsoft.com/en-us/graph/sdks/choose-authentication-providers?tabs=CS) 
1. [Microsoft Graph user: sendMail](https://learn.microsoft.com/en-us/graph/api/user-sendmail?view=graph-rest-1.0&tabs=csharp) 
1. [Microsoft Graph List Messages](https://learn.microsoft.com/en-us/graph/api/user-list-messages?view=graph-rest-1.0&tabs=csharp)

## Copyright and Ownership

All terms used are copyright to their original authors.