import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
import base64
from bs4 import BeautifulSoup
import traceback
import pandas as pd

SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

creds = None

if os.path.exists('token.json'):
    creds = Credentials.from_authorized_user_file('token.json', SCOPES)

if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)
    with open('token.json', 'w') as token:
        token.write(creds.to_json())

service = build('gmail', 'v1', credentials=creds)

df = pd.DataFrame([], columns=["Date", "Name", "Email", "State", "Phone"])
dict_list = []

inbox_label = service.users().labels().get(userId='me', id='INBOX').execute()
inbox_message_count = inbox_label['messagesTotal']
inbox_unread_count = inbox_label['messagesUnread']
print("Inbox message count:", inbox_message_count)
# print("Inbox unread count: ", inbox_unread_count)

results = service.users().messages().list(userId='me', labelIds=['INBOX'], maxResults=inbox_message_count).execute()
messages = results['messages']

for i in messages:
    id = i['id']
    message = service.users().messages().get(userId='me', id=id).execute()
    headers = message['payload']['headers']
    parts = message['payload']['parts']

    subject = ''
    for d in headers:
        if d['name'] == 'Subject':
            subject = d['value']
        if d['name'] == 'From':
            sender = d['value']

    if subject.startswith('Re: Forwarded Candidate: '):
        try:
            try:
                data = parts[0]['parts'][0]['body']['data']
            except:
                try:
                    data = parts[0]['body']['data']
                except:
                    traceback.print_exc()
                    continue

            data = data.replace("-", "+").replace("_", "/")
            decoded_data = base64.b64decode(data)
            soup = BeautifulSoup(decoded_data, "lxml")
            body = str(soup.body()[0])

            date = ''
            name_index = body.index('Name:\r\n&gt;') + len('Name:\r\n&gt;')
            name = body[name_index + 1:body.find('\r\n', name_index)]

            email_index = body.index('Email:', name_index) + len('Email:')
            email = body[email_index + 1:body.find('\r\n', email_index)]

            job_location_index = body.index('Job Location:', name_index) + len('Job Location:')
            state_index = body.index(', ', job_location_index) + len(', ')
            state = body[state_index:body.find('\r\n', state_index)]

            phone_index = body.index('Phone:', name_index) + len('Phone:')
            phone = body[phone_index + 1:body.find('\r\n', phone_index)]

            print(date, name, email, phone, state)
            dict_list.append({"Name": name, "Date": date, "Email": email, "Phone": phone, "State": state})

        except Exception:
            traceback.print_exc()

    elif subject == '[new_open_enrollment@app.getresponse.com] Subscription notification from getresponse':
        try:
            data = parts[0]['body']['data']
            data = data.replace("-", "+").replace("_", "/")
            decoded_data = base64.b64decode(data)
            soup = BeautifulSoup(decoded_data, "lxml")
            body = str(soup.body()[0])

            date_index = body.index('Timestamp: ') + len('Timestamp: ')
            date = body[date_index:body.index(' ', date_index)]

            name_index = body.index('Name: ') + len('Name: ')
            name = body[name_index:body.find('\n', name_index)]

            email_index = body.index('Email: ') + len('Email: ')
            email = body[email_index:body.index('\n', email_index)]

            state_index = body.index('state: ') + len('state: ')
            state = body[state_index:body.index('\n', state_index)]

            phone_index = body.index('maincontactnumber: ') + len('maincontactnumber: ')
            phone = body[phone_index:body.index('\n', phone_index)]

            print(date, name, email, phone, state)
            dict_list.append({"Name": name, "Date": date, "Email": email, "Phone": phone, "State": state})

        except Exception:
            traceback.print_exc()

    else:
        print("Subject unknown:", subject)
        continue


df = df.append(dict_list)
df.to_excel('sheet.xlsx', index=False)