Deployment Instructions:
1. Host all these files in your GitHub Static Web App (you already have it).
2. Upload the manifest.xml to your Office 365 Tenant -> Integrated Apps -> Upload Custom Add-in.
3. Open Outlook -> Open any email -> Add-in will automatically try to delete attachments.

Requirements:
- App Registration configured with API permissions (Mail.ReadWrite).
- Static Web App should allow CORS (default is OK).
- Silent SSO is triggered automatically using Office.js auth.
- 
