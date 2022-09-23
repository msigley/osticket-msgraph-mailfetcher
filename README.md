# Microsoft Graph API Mailfetcher Plugin for OSTicket
Fetches mail from Microsoft 365 using the Microsoft Graph API. This was written eliminate the use of IMAP with pulling mail from 365 accounts as Microsoft has been activing trying to eliminate it for years.

# Installation
1. Copy the `/msgraph/` folder to the `/include/plugins/` in your OSTicket installation.
2. Install and enable the plugin as outlined here: https://docs.osticket.com/en/latest/Admin/Manage/Plugins.html

# Configuration
1. Register a new app for OSTicket in Azure AD as outlined here: https://learn.microsoft.com/en-us/graph/auth-v2-service?context=graph%2Fapi%2F1.0&view=graph-rest-1.0
2. Durring registration select "Accounts in this organizational directory only" and leave Redirect URI blank.
3. Under "API Permissions", add the following permissions to the app under Microsoft Graph API > Application permissions > Mail: `Mail.ReadWrite, User.Read.All`
4. Provide Admin consent for these permissions.
5. Under "Certificates & secrets", add a new client secret and note the value of it for later. When asked for an expiration, I would suggest 24 months if you don't want to have to update this value in OSTicket often.
6. Under "Overview", note the value of "Application (client) ID" down for later.
7. In your OSTicket Admin Panel, go to Manage > Plugins.
8. Click on "Microsoft Graph API Mailfetcher".
9. Provide values for the settings here. If you don't know your Tenant Id, follow the directions here: https://learn.microsoft.com/en-us/azure/active-directory/fundamentals/active-directory-how-to-find-tenant
10. Click "Save Changes". A test connection will be made to validate the settings here and you will recieve errors if there are issues to resolve.
11. If you haven't done so, I would highly recommend you setup external cron handling: https://docs.osticket.com/en/latest/Developer%20Documentation/API/Tasks.html
