"""
Main script to read Outlook 365 mailbox using app-only authentication
Tracks last email and only retrieves new ones on subsequent runs
Runs in a continuous loop checking for new emails at configurable intervals
Minimal output - only shows new emails as they arrive
Optional SMTP forwarding for new emails
"""

import os
import smtplib
import time
from typing import List, Optional

from dotenv import load_dotenv

from mail_reader import O365MailReader

# Load environment variables from .env file
load_dotenv()


def _get_bool(key: str, default: bool = False) -> bool:
    """Parse boolean from environment variable"""
    value = os.getenv(key, str(default)).lower()
    return value in ("true", "1", "yes", "on")


def _get_list(key: str, default: Optional[List[str]] = None) -> List[str]:
    """Parse comma-separated list from environment variable"""
    value = os.getenv(key, "")
    if not value:
        return default or []
    return [item.strip() for item in value.split(",") if item.strip()]


# Azure AD App Registration Credentials
CLIENT_ID = os.getenv("CLIENT_ID", "")
TENANT_ID = os.getenv("TENANT_ID", "")
CLIENT_SECRET = os.getenv("CLIENT_SECRET", "")
MAILBOX_EMAIL = os.getenv("MAILBOX_EMAIL", "")

# Loop configuration
LOOP_DELAY_SECONDS = int(os.getenv("LOOP_DELAY_SECONDS", "5"))
ENABLE_CONTINUOUS_LOOP = _get_bool("ENABLE_CONTINUOUS_LOOP", True)

# SMTP forwarding (optional)
ENABLE_SMTP_FORWARD = _get_bool("ENABLE_SMTP_FORWARD", False)
SMTP_HOST = os.getenv("SMTP_HOST", "")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
SMTP_USERNAME = os.getenv("SMTP_USERNAME", "")
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD", "")
SMTP_USE_TLS = _get_bool("SMTP_USE_TLS", True)
SMTP_FROM = os.getenv("SMTP_FROM", "")
SMTP_TO = _get_list("SMTP_TO")


def _normalize_recipients(value) -> List[str]:
    if isinstance(value, str):
        return [addr.strip() for addr in value.split(",") if addr.strip()]
    return [addr for addr in value if addr]


def forward_message(reader: O365MailReader, msg: dict):
    """Forward a single message via SMTP"""
    recipients = _normalize_recipients(SMTP_TO)
    if not ENABLE_SMTP_FORWARD or not SMTP_HOST or not recipients:
        return

    message_id = msg.get("id")
    if not message_id:
        return

    # Get the raw MIME content of the message (preserves everything as-is)
    try:
        mime_content = reader.get_message_mime(message_id)
    except Exception as mime_err:
        print(f"❌ Failed to retrieve message MIME: {mime_err}")
        return

    try:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=30) as smtp:
            if SMTP_USE_TLS:
                smtp.starttls()
            if SMTP_USERNAME:
                smtp.login(SMTP_USERNAME, SMTP_PASSWORD)
            smtp.sendmail(SMTP_FROM, recipients, mime_content)
            print(f"✅ Forwarded via SMTP")
    except Exception as smtp_err:
        print(f"❌ SMTP forward failed: {smtp_err}")


def check_for_new_emails(reader: O365MailReader) -> List[dict]:
    """Check for new emails, display them, and forward each one immediately"""
    try:
        # Get access token silently
        if not reader.access_token:
            reader.get_access_token()

        # Get only new messages since last run
        messages = reader.get_new_messages(folder="inbox", limit=50)

        # Process each message: display info and forward immediately
        for msg in messages:
            sender = msg.get("from", {}).get("emailAddress", {}).get("address", "Unknown")
            to_addr = msg.get("toRecipients", [{}])[0].get("emailAddress", {}).get("address", "Unknown") if msg.get("toRecipients") else "Unknown"
            subject = msg.get("subject", "[No Subject]")
            has_attachments = msg.get("hasAttachments", False)

            print(f"\n{'='*80}")
            print(f"📧 From: {sender}")
            print(f"   To: {to_addr}")
            print(f"   Subject: {subject}")
            if has_attachments:
                message_id = msg.get("id")
                if message_id:
                    try:
                        attachments = reader.get_attachments(message_id)
                        att_names = [att.get("name", "unknown") for att in attachments]
                        print(f"   Attachments ({len(attachments)}): {', '.join(att_names)}")
                    except Exception:
                        print(f"   Has attachments: Yes")
                else:
                    print(f"   Has attachments: Yes")

            # Forward this message immediately
            forward_message(reader, msg)

        return messages

    except Exception as e:
        print(f"❌ Error: {e}")
        return []


def main():
    """Main function to run continuous mail checking"""

    print("🔔 Outlook 365 Email Monitor Started - Press Ctrl+C to stop\n")

    run_count = 0
    new_emails_count = 0
    reader = O365MailReader(CLIENT_ID, TENANT_ID, CLIENT_SECRET, MAILBOX_EMAIL)

    try:
        while True:
            run_count += 1

            # Check for new emails (forwards happen inside this function now)
            messages = check_for_new_emails(reader)
            if messages:
                new_emails_count += len(messages)

            # Exit if not in continuous loop mode
            if not ENABLE_CONTINUOUS_LOOP:
                break

            # Wait before next check
            time.sleep(LOOP_DELAY_SECONDS)

    except KeyboardInterrupt:
        print(f"\n\n✓ Monitor stopped")
        print(f"  Total checks: {run_count}")
        print(f"  Total new emails: {new_emails_count}")

    except Exception as e:
        print(f"\n❌ Fatal error: {e}")
        return 1

    return 0


if __name__ == "__main__":
    exit(main())
