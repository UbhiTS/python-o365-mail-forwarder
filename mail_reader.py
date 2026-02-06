"""
Outlook 365 Mail Reader using App-only Authentication
Reads mailbox using Azure AD app-only auth with Microsoft Graph API
"""

import requests
import json
import os
from datetime import datetime
from typing import List, Dict, Any


class O365MailReader:
    """Class to handle reading Outlook 365 mailbox using app-only auth"""
    
    def __init__(self, client_id: str, tenant_id: str, client_secret: str, mailbox_email: str):
        self.client_id = client_id
        self.tenant_id = tenant_id
        self.client_secret = client_secret
        self.mailbox_email = mailbox_email
        self.token_endpoint = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
        self.scopes = ["https://graph.microsoft.com/.default"]
        self.access_token = None
        self.tracking_file = ".last_email.json"
        self.silent_mode = True  # Suppress verbose output
        
    def get_access_token(self) -> str:
        """
        Get an access token using client credentials flow (app-only auth)
        
        Returns:
            str: Access token for Graph API
        """
        payload = {
            "client_id": self.client_id,
            "scope": " ".join(self.scopes),
            "client_secret": self.client_secret,
            "grant_type": "client_credentials"
        }
        
        try:
            response = requests.post(self.token_endpoint, data=payload)
            response.raise_for_status()
            
            token_data = response.json()
            self.access_token = token_data.get("access_token")
            
            if not self.access_token:
                raise Exception("Failed to retrieve access token")
            
            if not self.silent_mode:
                print(f"✓ Successfully obtained access token")
            return self.access_token
            
        except requests.exceptions.RequestException as e:
            print(f"✗ Error obtaining access token: {e}")
            if hasattr(e, 'response') and e.response is not None:
                print(f"Response: {e.response.text}")
            raise
    
    def get_attachments(self, message_id: str) -> List[Dict[str, Any]]:
        """
        Get attachments from a specific message
        
        Args:
            message_id: The message ID
            
        Returns:
            List of attachment dictionaries
        """
        if not self.access_token:
            self.get_access_token()
        
        url = f"https://graph.microsoft.com/v1.0/users/{self.mailbox_email}/messages/{message_id}/attachments"
        
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        
        try:
            response = requests.get(url, headers=headers)
            response.raise_for_status()
            
            attachments = response.json().get("value", [])
            return attachments
            
        except requests.exceptions.RequestException as e:
            print(f"✗ Error retrieving attachments: {e}")
            raise

    def get_message_mime(self, message_id: str) -> bytes:
        """
        Get the raw MIME content of a message (RFC 822 format)
        
        Args:
            message_id: The message ID
            
        Returns:
            bytes: Raw MIME content of the message
        """
        if not self.access_token:
            self.get_access_token()
        
        url = f"https://graph.microsoft.com/v1.0/users/{self.mailbox_email}/messages/{message_id}/$value"
        
        headers = {
            "Authorization": f"Bearer {self.access_token}"
        }
        
        try:
            response = requests.get(url, headers=headers)
            response.raise_for_status()
            return response.content
            
        except requests.exceptions.RequestException as e:
            print(f"✗ Error retrieving message MIME: {e}")
            raise
    
    def load_last_email_timestamp(self) -> tuple:
        """
        Load the timestamp and ID of the last processed email from disk
        
        Returns:
            tuple: (timestamp string, message_id) of last email, or (None, None) if no tracking file exists
        """
        if not os.path.exists(self.tracking_file):
            return None, None
        
        try:
            with open(self.tracking_file, 'r') as f:
                data = json.load(f)
                return data.get("last_received_datetime"), data.get("last_message_id")
        except (json.JSONDecodeError, IOError) as e:
            print(f"Warning: Could not read tracking file: {e}")
            return None, None
    
    def save_last_email_timestamp(self, message: Dict[str, Any]):
        """
        Save the timestamp of the latest processed email to disk
        
        Args:
            message: The latest message dictionary
        """
        try:
            data = {
                "last_received_datetime": message.get("receivedDateTime"),
                "last_message_id": message.get("id"),
                "last_subject": message.get("subject"),
                "last_from": message.get("from", {}).get("emailAddress", {}).get("address"),
                "last_updated": datetime.now().isoformat()
            }
            
            with open(self.tracking_file, 'w') as f:
                json.dump(data, f, indent=2)
            
            if not self.silent_mode:
                print(f"✓ Tracking file updated with latest email")
            
        except IOError as e:
            print(f"Warning: Could not save tracking file: {e}")
    
    def get_new_messages(self, folder: str = "inbox", limit: int = 50) -> List[Dict[str, Any]]:
        """
        Get only new messages since the last run.
        
        This method checks the tracking file to find the last processed email
        and retrieves only newer messages using Graph API filtering.
        
        Args:
            folder: Folder name (inbox, drafts, sentitems, deleteditems, etc.)
            limit: Maximum number of new messages to retrieve
            
        Returns:
            List of new message dictionaries
        """
        if not self.access_token:
            self.get_access_token()
        
        # Load the last email timestamp and ID
        last_timestamp, last_message_id = self.load_last_email_timestamp()
        
        # Graph API endpoint
        url = f"https://graph.microsoft.com/v1.0/users/{self.mailbox_email}/mailFolders/{folder}/messages"
        
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        
        # Build filter for new messages
        params = {
            "$top": limit,
            "$select": "id,subject,from,toRecipients,receivedDateTime,sender,hasAttachments,isRead,bodyPreview",
            "$orderby": "receivedDateTime desc"
        }
        
        # If we have a previous timestamp, filter to get only newer messages
        if last_timestamp:
            # Use OData filter to get messages received strictly after the last one
            params["$filter"] = f"receivedDateTime gt {last_timestamp}"
            if not self.silent_mode:
                print(f"Fetching new messages since: {last_timestamp}")
        else:
            if not self.silent_mode:
                print("First run - fetching recent messages")
        
        try:
            response = requests.get(url, headers=headers, params=params)
            response.raise_for_status()
            
            messages = response.json().get("value", [])
            
            # Filter out the last message we already processed (in case of timestamp collisions)
            if last_message_id and messages:
                messages = [msg for msg in messages if msg.get("id") != last_message_id]
            
            if messages:
                if not self.silent_mode:
                    print(f"✓ Found {len(messages)} new messages")
                # Update tracking file with the newest message
                self.save_last_email_timestamp(messages[0])
            else:
                if not self.silent_mode:
                    print("✓ No new messages")
            
            return messages
            
        except requests.exceptions.RequestException as e:
            print(f"✗ Error retrieving messages: {e}")
            if hasattr(e, 'response') and e.response is not None:
                print(f"Response: {e.response.text}")
            raise
