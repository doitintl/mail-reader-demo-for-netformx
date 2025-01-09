import msal
import requests
import os

# Configuration
client_id = os.environ.get("CLIENT_ID") # Application (client) ID
client_secret = os.environ.get("CLIENT_SECRET") # Client secret
tenant_id = os.environ.get("TENANT_ID") # Directory (tenant) ID
user_email = os.environ.get("USER_EMAIL") # Email address of the mailbox to read

# Microsoft Graph endpoints
authority = f"https://login.microsoftonline.com/{tenant_id}"
scopes = ["https://graph.microsoft.com/.default"]  # For application permissions
graph_endpoint = f"https://graph.microsoft.com/v1.0/users/{user_email}/messages"

# Authentication using MSAL
def get_access_token():
    app = msal.ConfidentialClientApplication(client_id, authority=authority, client_credential=client_secret)
    result = app.acquire_token_for_client(scopes)
    if "access_token" in result:
        return result["access_token"]
    else:
        raise Exception(f"Could not acquire token: {result.get('error_description', result)}")

# List email subjects
def list_email_subjects():
    token = get_access_token()
    headers = {"Authorization": f"Bearer {token}"}
    
    response = requests.get(graph_endpoint, headers=headers)
    if response.status_code == 200:
        emails = response.json().get("value", [])
        for email in emails:
            print(f"Subject: {email.get('subject', 'No Subject')}")
    else:
        print(f"Error: {response.status_code}, {response.text}")

if __name__ == "__main__":
    list_email_subjects()