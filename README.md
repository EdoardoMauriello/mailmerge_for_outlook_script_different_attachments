# mailmerge_for_outlook_ps_script-single_attachement
# Outlook Email Sender with a single attachment

This PowerShell script automates the process of sending personalized emails with attachments using Microsoft Outlook. It's useful for sending bulk emails with customized content and a single attachment.

## How to Use

The PowerShell script to uses an Excel sheet for the receiving email addresses and add names to the email. To achieve this, you can use the Import-Excel cmdlet from the ImportExcel module if you have it installed. To install, run the below command as administrator in PowerShell:
Install-Module -Name ImportExcel

1. Ensure you have Microsoft Outlook installed and configured on your system.

2. Prepare your recipient data in an Excel file. The Excel file should contain at least two columns: one for the recipient's name and another for their email address.

3. Modify the script variables:
   - `$excelFilePath`: Set the path to your Excel file with recipient data.
   - `$attachmentPath`: Update this to the path of the attachment you want to send.

4. Run the script using PowerShell.

5. The script will loop through your Excel data and send personalized emails to each recipient with the specified attachment.

6. After sending 10 emails, the script will pause for 1 minute to avoid overloading the email server. You can adjust this interval as needed.

7. The script will display a message once all emails have been sent.
