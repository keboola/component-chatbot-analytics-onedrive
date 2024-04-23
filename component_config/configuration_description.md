Accepts following parameters:

- Sharepoint username - example: data_ca@keboolads.onmicrosoft.com
- Sharepoint password
- Authority - example: https://login.microsoftonline.com/{tenant-id}/

- Office 365 Client ID - example: 25d65042-xxxx-xxxx-xxxx-xxxxxxxx
- Office 365 Client Secret
- Microsoft Tenant ID - example: 6f4e5157-xxxx-xxxx-xxxx-xxxxxxxxx
- Hostname - example: keboolads.sharepoint.com
- Url - example: sites/chatbot-files

Row configuration:

- Main Folder Path - example: chatbot/subfolder/Maintenance/
- Operation Type - download/upload
- Date From - Enter relative day (X days ago) or specific date in YYYY-MM-DD format.
- Date To [Exclusive] - Relative date (X days ago) or specific date in YYYY-MM-DD format.
- Filter files by date created - When set to true, this option creates a folder for each day within the selected period and uploads files for that specific day. When set to false, it uploads all files from the input mapping regardless of the date created. Use this setting only for periods of one day to prevent duplicity.
- Folder Suffix - String that will be appended to folder names for both download and upload.
