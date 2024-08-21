# Script for Managing App Secret & Certificate Expirations


## Variables:


*These variables are can be hard-coded or set in automation account variables, as I've done here*
- `$AppID`: Retrieved from shared resources, used to identify the app in Azure AD.
- `$TenantID`: Tenant ID for the Azure AD instance.
- `$AppSecret`: Secret key for authenticating the app.


*Set within first few lines of the script, change as needed*
- `$expirationDays`: If secret/certificate is within this many days of expiring, it will be added to the email.
- `$emailSender`: The email address that will send the notifications.
- `$emailTo`: Recipient email addresses for the notifications. Add multiple by separating emails with commas (ie. ```"kaiden.ong000@gmail.com", "kaiden.ong000@gmail.com"```)


## Functions:
- `Connect-MSGraphAPI`: Connects to Microsoft Graph API and retrieves an access token.
- `Get-MSGraphRequest`: Retrieves all app registrations and handles pagination.
- `Send-MSGraphEmail`: Sends an email with the provided parameters through Graph API.


## Main Logic:
The script uses the functions to connect to Graph API, retrieve app registrations, check for expiring secrets/certificates, and send them as tables, notifying the specified email addresses.


## Usage:
- Intended to be run on a scheduled basis, this can be done under the automation account → Resources → Schedules.


## Debugging Steps:
- Check settings under App Registration
  - Must have API permissions for Application.Read.All & Mail.Send
  - Make sure secret key is not expired
- Run the runbook and ensure output has no errors and returns as expected
  - Common problems include invalid sender or recipients