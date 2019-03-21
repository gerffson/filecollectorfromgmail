from __future__ import print_function
import pickle
import os.path
import sys
import base64
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from apiclient import errors

SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

def main():
    qry = sys.argv[1] + ' has:attachment'
    path  = sys.argv[2]
    service = build('gmail', 'v1', credentials=GetCredentials())
    messages = GetMessagesByQuery(service, 'me', query= qry)
    if not messages:
        print('No messages found.')
    else:
        print('Messages:')
        for msg in messages:
            GetAttachments(service, 'me', msg['id'], path)


def GetCredentials():
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server()
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    return creds


def GetMessagesByQuery(service, user_id, query=''):
  try:
    response = service.users().messages().list(userId=user_id,q=query).execute()
    messages = []
    if 'messages' in response:
      messages.extend(response['messages'])

    while 'nextPageToken' in response:
      page_token = response['nextPageToken']
      response = service.users().messages().list(userId=user_id, q=query,pageToken=page_token).execute()
      messages.extend(response['messages'])
    
    return messages
  except Exception as e:
    print('An error occurred: ListMessagesMatchingQuery: ' + str(e))


def GetAttachments(service, user_id, msg_id, store_dir):
  
  try:
    message = service.users().messages().get(userId=user_id, id=msg_id).execute()
    for part in message['payload']['parts']:
      file_name = part['filename']
      if file_name.lower().endswith(('.pdf', '.jpg', '.jpeg')):
        newvar = part['body']
        if 'attachmentId' in newvar:
          att_id = newvar['attachmentId']        
          GetFileAttached(service, user_id, msg_id, att_id, store_dir, file_name)
  except Exception as e:
    print('An error occurred : GetAttachments : ' + str(e))


def GetFileAttached(service, user_id, msg_id, att_id, store_dir, file_name):
  print ('[user_id : ' + user_id + '] [msg_id : ' + msg_id + '] [file : ' + store_dir + file_name + ']')
  try:
    att = service.users().messages().attachments().get(userId=user_id, messageId=msg_id, id=att_id).execute()    
    file_data = base64.urlsafe_b64decode(att['data'].encode('UTF-8'))
    path = ''.join([store_dir, file_name])
    f = open(path, 'wb')
    f.write(file_data)
    f.close()
  except Exception as e:
    print('An error occurred : GetFileAttached : ' + str(e))



if __name__ == '__main__':
    main()