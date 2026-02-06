# Outlook 365 Mail Reader & Forwarder

This Python program reads messages from an Outlook 365 mailbox using app-only authentication (Azure AD Client Credentials flow) and optionally forwards them via SMTP.

## Features

- ✓ App-only authentication (no user interaction required)
- ✓ Continuous email monitoring with configurable polling interval
- ✓ Tracks last processed email to only show new messages
- ✓ SMTP forwarding with full message preservation (headers, body, attachments)
- ✓ Raw MIME forwarding - emails are forwarded exactly as received
- ✓ Attachment support - displays attachment names in console
- ✓ Minimal console output showing From, To, Subject, and Attachments
- ✓ Configurable via `.env` environment file

## Prerequisites

- Python 3.7+
- Microsoft 365 account with admin access
- Azure AD app registration with:
  - Client ID
  - Tenant ID
  - Client Secret
  - Proper API permissions (Mail.Read on behalf of the mailbox)

## Setup

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

### 2. Configure Environment Variables

Copy the example environment file and edit with your credentials:

```bash
cp .env.example .env
```

Edit `.env` with your Azure AD app registration details:

```env
# Azure AD App Registration Credentials
CLIENT_ID=your-client-id
TENANT_ID=your-tenant-id
CLIENT_SECRET=your-client-secret

# Mailbox to access
MAILBOX_EMAIL=admin@example.com

# Loop configuration
LOOP_DELAY_SECONDS=5
ENABLE_CONTINUOUS_LOOP=true

# SMTP forwarding (optional)
ENABLE_SMTP_FORWARD=false
SMTP_HOST=smtp.example.com
SMTP_PORT=587
SMTP_USERNAME=your-username
SMTP_PASSWORD=your-password
SMTP_USE_TLS=true
SMTP_FROM=forwarder@example.com
SMTP_TO=recipient1@example.com,recipient2@example.com
```

> ⚠️ **Never commit `.env` to version control.** The `.env` file contains sensitive credentials and should be added to `.gitignore`.

### 3. Grant API Permissions

In Azure Portal, ensure your app has the following permissions:
- **Mail.Read** (Application permission for reading mail)
- **User.Read** (Application permission for reading user info)

Grant admin consent for these permissions.

### 4. Configure App-Only Auth in Azure AD

The app uses the **Client Credentials OAuth 2.0 flow**, which requires:
1. Application is registered as a multi-tenant or single-tenant app
2. Proper permissions are granted by admin
3. Mailbox to be accessed must have the app permissions applied (e.g., via PowerShell)

**PowerShell setup (if needed):**
```powershell
# Connect to Exchange Online
Connect-ExchangeOnline

# Grant permission to the app to read the mailbox
Add-RecipientPermission -Identity "admin@ubhims.com" -Trustee "application-name-from-azure" -AccessRights ReadPermission
```

## Usage

### Running the Monitor

```bash
python main.py
```

This will:
1. Authenticate using app-only credentials
2. Start monitoring the inbox for new emails
3. Display new emails with From, To, Subject, and Attachments
4. Optionally forward emails via SMTP (if configured)
5. Continue polling at the configured interval (default: 5 seconds)

Press `Ctrl+C` to stop the monitor.

### Console Output Example

```
🔔 Outlook 365 Email Monitor Started - Press Ctrl+C to stop

================================================================================
📧 From: sender@example.com
   To: recipient@example.com
   Subject: Important Document
   📎 Attachments (2): document.pdf, image.png
   ✅ Forwarded via SMTP
================================================================================
```

### SMTP Forwarding

When `ENABLE_SMTP_FORWARD = True`, emails are forwarded using raw MIME content, which means:
- Original headers are preserved
- HTML/plain text body is forwarded as-is
- All attachments are included exactly as received
- No reconstruction or re-encoding of the message

### Advanced Usage

```python
from mail_reader import O365MailReader

# Initialize reader with credentials
reader = O365MailReader(
    client_id="your-client-id",
    tenant_id="your-tenant-id",
    client_secret="your-client-secret",
    mailbox_email="mailbox@example.com"
)

# Get access token
reader.get_access_token()

# Get new messages since last check
messages = reader.get_new_messages(folder="inbox", limit=50)

# Get message details
if messages:
    msg_id = messages[0]['id']
    attachments = reader.get_attachments(msg_id)
    mime_content = reader.get_message_mime(msg_id)  # Raw RFC 822 content
```

## API Folders

Available folders in Graph API:
- `inbox` - Inbox folder
- `drafts` - Drafts folder
- `sentitems` - Sent Items folder
- `deleteditems` - Deleted Items folder
- `archive` - Archive folder

## Message Fields

The `get_new_messages()` method returns messages with:
- `id` - Message ID
- `subject` - Message subject
- `from` - Sender information
- `toRecipients` - Recipient information
- `receivedDateTime` - Received date and time
- `hasAttachments` - Boolean indicating attachments
- `isRead` - Read status
- `bodyPreview` - First 256 characters of body

## Mail Reader Methods

| Method | Description |
|--------|-------------|
| `get_access_token()` | Authenticate and get access token |
| `get_new_messages(folder, limit)` | Get only new messages since last check |
| `get_attachments(message_id)` | Get attachment metadata and content |
| `get_message_mime(message_id)` | Get raw MIME content (RFC 822) |

## Error Handling

The program includes error handling for:
- Authentication failures
- Network errors
- Invalid mailbox
- API permission issues
- Missing messages

Check the console output for detailed error messages.

## Troubleshooting

### "InvalidAuthenticationToken" Error
- Verify CLIENT_ID, TENANT_ID, and CLIENT_SECRET are correct
- Ensure admin consent has been granted in Azure Portal
- Check that credentials haven't expired

### "Access Denied" Error
- Verify the app has proper permissions in Azure AD
- Ensure admin consent was granted
- Check that the mailbox exists and is accessible

### "Mail.Read" Permission Not Found
- Go to Azure Portal > App Registration > API Permissions
- Click "Add a Permission" > Microsoft Graph > Application Permissions
- Search for "Mail.Read" and grant it
- Click "Grant admin consent for [Organization]"

## Security Notes

⚠️ **Important Security Considerations:**

1. **Never commit `.env` to version control** - Add `.env` to your `.gitignore`:
   ```
   .env
   ```

2. **Rotate client secrets regularly** in Azure Portal

3. **Use Azure Key Vault** in production for storing secrets

4. **Audit API access logs** in Microsoft 365 admin center

5. **Principle of least privilege** - Only grant necessary permissions

## References

- [Microsoft Graph API Documentation](https://docs.microsoft.com/en-us/graph/overview)
- [OAuth 2.0 Client Credentials Flow](https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-client-creds-grant-flow)
- [Outlook Mail API](https://docs.microsoft.com/en-us/graph/api/resources/message)
- [Azure AD App Registration](https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app)

## Disclaimer

⚠️ **USE AT YOUR OWN RISK**

This software is provided "as-is" without warranty of any kind, express or implied. The author(s) shall not be held liable for any damages, data loss, security breaches, or other issues arising from the use of this software.

**Important notices:**

- This project is **NOT authorized, endorsed, or vetted by Microsoft Corporation**
- This is an independent, community-developed tool that uses publicly available Microsoft Graph APIs
- You are solely responsible for ensuring compliance with your organization's policies, Microsoft's Terms of Service, and applicable laws and regulations
- The author(s) make no guarantees about the reliability, security, or suitability of this code for any purpose
- Before using in production, thoroughly review and test the code in a safe environment
- You are responsible for securing your credentials and protecting sensitive data

By using this software, you acknowledge that you understand these risks and accept full responsibility for any consequences.

## License

This project is provided as-is for educational purposes.
