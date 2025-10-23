"""
Configuration file for SharePoint authentication
Store your credentials here securely
"""

# SharePoint Configuration
SHAREPOINT_URL = "https://unitechn.sharepoint.com"
SHAREPOINT_SITE_URL = "https://unitechn.sharepoint.com/sites/TutoriasUNITEC2/"

# Interactive Authentication (Recommended - Modern and Secure)
# Uses Microsoft Graph PowerShell client ID for interactive authentication
INTERACTIVE_AUTH_ENABLED = True
INTERACTIVE_CLIENT_ID = "04b07795-8ddb-461a-bbee-02f9e1bf7b46"  # Microsoft Graph PowerShell
INTERACTIVE_TENANT_ID = "common"  # Use 'common' for multi-tenant
INTERACTIVE_SCOPE = ["https://graph.microsoft.com/.default"]

# Basic Authentication (Deprecated - may not work due to Microsoft's security policies)
SHAREPOINT_USERNAME = "your_username@unitec.edu"  # Replace with your actual username
SHAREPOINT_PASSWORD = "your_password"  # Replace with your actual password

# App Registration (Alternative for production)
SHAREPOINT_CLIENT_ID = "your_client_id"
SHAREPOINT_CLIENT_SECRET = "your_client_secret"
SHAREPOINT_TENANT_ID = "your_tenant_id"

# Alternative: Certificate-based authentication
CERTIFICATE_PATH = "path/to/your/certificate.pem"
CERTIFICATE_PASSWORD = "certificate_password"
