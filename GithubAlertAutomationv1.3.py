import imaplib
import email
import os.path

import pandas as pd
from bs4 import BeautifulSoup
from urllib.parse import urlparse, parse_qs

from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
import openpyxl
import datetime

# Set up IMAP server details
imap_server = 'imap.gmail.com'
username = 'example@gmail.com'
password = 'APP PASSWORD' # This must be an app password, this can be enabled at the 2fa page on google

def Find_Links(subject):
    # Connect to the IMAP server
    server = imaplib.IMAP4_SSL(imap_server)
    server.login(username, password)

    # Select the mailbox you want to access (e.g., 'INBOX')
    server.select('INBOX')

    # Search for emails with the desired subject line
    status, response = server.search(None, subject)

    # Get the latest email ID
    latest_email_id = response[0].split()[-1]

    # Fetch the email content
    status, email_data = server.fetch(latest_email_id, '(RFC822)')
    raw_email = email_data[0][1].decode('utf-8')

    # Parse the raw email content
    email_message = email.message_from_string(raw_email)

    # Extract the subject of the email
    subject = email_message['Subject']
    substring_to_remove = "Google Alert - "
    modified_string = subject.replace(substring_to_remove, "")




    # Extract the date of the email
    email_date = email_message['Date']
    parsed_email_date = email.utils.parsedate_to_datetime(email_date)
    formatted_email_date = parsed_email_date.strftime("%m-%d-%Y")





    # Check if the email is a Google Alert email
    if 'google alert' in subject.lower():
        # Extract all the links from the email
        links = []
        for part in email_message.walk():
            if part.get_content_type() == 'text/html':
                soup = BeautifulSoup(part.get_payload(decode=True), 'html.parser')
                links.extend(soup.find_all('a'))

        # Extract the URLs from the links
        urls = [link.get('href') for link in links if link.get('href')]



        i = 0
        urls2 = []
        for url in urls:
            urls2.append(0)
            parsed_url = urlparse(url)  # Parse the URL
            query_params = parse_qs(parsed_url.query)
            middle_link = query_params.get('url', [''])[0]

            if not middle_link:
                # Try extracting the middle link from a different URL format
                path_segments = parsed_url.path.split("/")
                if len(path_segments) > 2:
                    middle_link = "/".join(path_segments[2:])
            urls2[i] = middle_link
            i += 1

        urls[0] = 0
        urls[-1] = 0
        urls[-2] = 0
        urls[-3] = 0
        urls[-4] = 0
        urls[-5] = 0
        urls[-6] = 0

        urls2[0] = 0
        urls2[-1] = 0
        urls2[-2] = 0
        urls2[-3] = 0
        urls2[-4] = 0
        urls2[-5] = 0
        urls2[-6] = 0


        # Create a DataFrame with the URLs
        df = pd.DataFrame({'URLS': urls})

        # Add a column with the subject name
        df['Subject'] = modified_string
        df['Formatted Urls'] = urls2
        df['Date'] = formatted_email_date
        negative_indices = [-1,-2,-3,-4,-5,-6]
        positive_indices = [len(df) + i if i < 0 else i for i in negative_indices]
        df = df.drop(positive_indices)
        df = df.drop(0)
        if 'Formatted Urls' in df.columns:
            for i in range(len(df)+1):
                if i in df.index and (
                        df['Formatted Urls'][i].lower() == "feedback" or df['Formatted Urls'][i].lower() == "share" or df['Formatted Urls'][i].lower() == "story"):
                    df = df.drop(i)

        # Note: Make sure to handle the remaining code accordingly after the loop

        # Save the DataFrame to an Excel file
        df = df.reindex(['Date', 'Subject', 'Formatted Urls'], axis=1)
        excel_file = 'links.xlsx'
        if (os.path.exists(excel_file)):
            df_read = pd.read_excel(excel_file)
            df_final = pd.concat([df_read, df], ignore_index=True)
            df_final.to_excel(excel_file, index=False)
        else:
            df.to_excel(excel_file, index=False)

    # Close the server connection
    server.close()
    server.logout()

Find_Links('SUBJECT "Google Alert - college admission"')
Find_Links('SUBJECT "Google Alert - financial aid"')
Find_Links('SUBJECT "Google Alert - internship"')
Find_Links('SUBJECT "Google Alert - volunteer"')
Find_Links('SUBJECT "Google Alert - scholarship"')

