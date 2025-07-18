# Automated Lead Tracker for Outlook

This script automatically processes new emails, uses AI to identify meeting requests, creates calendar events, and logs all lead interactions in an Excel spreadsheet.

## Setup Instructions

### Prerequiste

- You must have Python 3.8 or newer installed on your computer.

### 

- Required files:
    - get_emails.py
    - README.md
    - requirements.txt

### 1. Configure Credentials
- Open the `get_emails.py` file in a text editor.
- Go to the `Configuration` section at the top of the file.
- Replace the placeholder values for `CLIENT_ID`, `TENANT_ID`, and `GEMINI_API_KEY` with your own credentials. Save the file.
- In order to get this values you must do the following:
    How to Get Microsoft Credentials
        -> Sign in to the Azure Portal and navigate to Microsoft Entra ID > App registrations.
        -> Click + New registration. Give it a name and select the "Accounts in any organizational directory... and personal Microsoft accounts" option, then click Register.
        -> On the Overview page, copy the Application (client) ID and the Directory (tenant) ID.
        -> Go to the Authentication tab, enable "Allow public client flows," and click Save.
        -> Go to the API permissions tab and add the following Delegated permissions for Microsoft Graph:
            -> Calendars.ReadWrite
            -> Mail.Read
            -> Files.ReadWrite
            -> User.Read

### 2. Install Dependencies
- Open your terminal and run the following command to install the required Python libraries:
  ```bash
  pip install -r requirements.txt

### 3. Run the Application 
- Once setup is complete, run the script from your terminal:
    -python get_emails.py
-The first time you run it, you will be prompted to log in to your Microsoft account in your browser to grant the script permission.

### 4. To Keep it Running in the background
- On Linux/macOS: Use the command nohup python get_emails.py &
- On Windows: Run python get_emails.py in a terminal and minimize the window.