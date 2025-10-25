"""
================================== Introduction: Gmail Auto Mail Sender ==================================

This Python script automates sending personalized email messages to multiple recipients
listed in an Excel file. It connects securely to the Gmail API using OAuth 2.0 authentication
and sends each email individually to ensure reliable delivery and avoid spam detection.

---------------------------------------- Features --------------------------------------------------------

- Reads recipient addresses from an Excel file named "Email List.xlsx" with a column titled "Email".
  (You must create this Excel file before running the script. Include only valid email addresses.)
- Sends a prewritten message to each recipient via Gmail API.
- Automatically handles Gmail authentication using credentials.json and token.pickle files.
- Tracks failed deliveries and updates the Excel file to keep only emails that failed to send.
- Encodes email messages in Base64, as required by Gmail API.
- Automates repetitive outreach tasks, saving time and reducing human error.

---------------------------------- Gmail API Setup Instructions ------------------------------------------

1️⃣ Go to https://console.cloud.google.com/ and sign in with your Gmail account.
2️⃣ Create a **new project** (e.g., "Python Mail Sender").
3️⃣ Enable the **Gmail API** from the Google Cloud Console.
4️⃣ Go to "API's & Services" → "Credentials" → "Create Credentials" → "OAuth Client ID".
5️⃣ Choose **Desktop App** and download the **credentials.json** file. (Important)
6️⃣ Save `credentials.json` in the same folder as this script.
7️⃣ When running the script for the first time, a browser window will open to authenticate
   and grant access. This will automatically generate a **token.pickle** file for future runs. (The "token" must be kept)

------------------------------------------ Important Notes -----------------------------------------------

- Keep `credentials.json` and `token.pickle` secure; do not share publicly.
- If 2FA is enabled, use an App Password or follow Gmail OAuth instructions to allow the script to send emails.
- Only valid email addresses in the Excel file will be processed. Empty or invalid entries are ignored.
- The script is ideal for job applications, recruitment, or professional outreach.

------------------------------------------------------------------------------------------------------------
Author: Bar Cohen
Language: Python 3.13
Libraries: pandas, google-api-python-client, google-auth, google-auth-oauthlib, email, base64, pickle
============================================================================================================
"""

import pandas as pd                                     # pandas library for handling Excel files and data frames
import base64                                           # used to encode messages for Gmail API
from email.mime.text import MIMEText                    # for creating plain text email messages
from googleapiclient.discovery import build             # builds the Gmail API service object
from google_auth_oauthlib.flow import InstalledAppFlow  # handles OAuth authentication flow
from google.auth.transport.requests import Request      # used to refresh expired tokens
import os.path                                          # file and path operations
import pickle                                           # used for saving/loading authentication tokens

# ---- 1. Read the email list from Excel ----
try:
    df = pd.read_excel("Email List.xlsx")               # make sure the Excel file exists and has a column named 'Email'
    emails = df["Email"].dropna().astype(str).str.strip().tolist()  # clean up the email list
except Exception as e:
    print(f"❌ Error reading Excel file: {e}")
    exit()

# ---- 2. Connect to Gmail API ----
SCOPES = ['https://www.googleapis.com/auth/gmail.send']  # permission to send emails only

creds = None
if os.path.exists('token.pickle'):                      # check if a token already exists
    with open('token.pickle', 'rb') as token:
        creds = pickle.load(token)                      # load the saved credentials

# if credentials are missing or invalid, perform the authentication flow
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())                        # refresh the token if it has expired
    else:
        # initial login using OAuth and credentials.json file
        flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES) # Make sure it is indeed found, and if not, pull it from Google.
        creds = flow.run_local_server(port=0)            # opens a browser window for user authentication
    # save the token for future use
    with open('token.pickle', 'wb') as token:
        pickle.dump(creds, token)

# build the Gmail service object (by the token)
service = build('gmail', 'v1', credentials=creds)

# ---- 3. Email content ----
subject = "Application for Relevant Opportunities - Bar Cohen"  # email subject line

body = """\
My name is Bar Cohen, and I hold a B.Sc. in Information Systems with a specialization in Data Science.  
Throughout my studies and projects, I gained hands-on experience in data analysis, machine learning, and software development, and I am now eager to take the next step in my professional career.  

I would be happy to know if there are any opportunities within your organization that could be relevant for my background and skills.  
I am available for a full-time position and can start on short notice.  

Thank you very much for your time and consideration.  
I would be glad to share my CV and provide further details if needed.  

Best regards,  
Bar Cohen  

LinkedIn: https://www.linkedin.com/in/bar--cohen-/  
GitHub: https://github.com/BarCohen-dot
"""

# ---- 4. Send emails safely ----
failed_emails = []  # list to collect emails that failed to send

for email in emails:  # iterate through all recipients
    try:
        # create a plain text email message
        message = MIMEText(body, "plain", "utf-8")
        message['to'] = email                                         # recipient address
        message['from'] = "me"                                        # "me" tells Gmail API to use the authenticated account
        message['subject'] = subject                                  # set the email subject

        # encode the message in Base64 as required by Gmail API
        raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
        message = {'raw': raw}                                        # format the message for sending

        # send the email
        sent = service.users().messages().send(userId="me", body=message).execute()
        print(f"✅ Sent to {email} (ID: {sent['id']})")

    except Exception as e:
        print(f"❌ Failed to send to {email}: {e}")
        failed_emails.append(email)  # add failed email to the list

# ---- 5. Update Excel file ----
try:
    # keep only failed emails for review
    remaining_df = df[df["Email"].isin(failed_emails)]
    remaining_df.to_excel("Email List.xlsx", index=False)  # overwrite the Excel file
    print("\n📄 Excel file updated: only failed emails remain.")
except Exception as e:
    print(f"⚠️ Could not update Excel file: {e}")

print("\n✅ All emails processed successfully!")
