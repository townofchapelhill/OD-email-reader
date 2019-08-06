# Dependencies
#from imapclient import IMAPClient
import imaplib
import base64
import os
import email
import pathlib
import secrets, filename_secrets

output_folder = pathlib.Path(filename_secrets.selfcheckStatistics)
searchQuery = '(SUBJECT "[Report]")'

# Function to login to IMAP
def login_and_write():

    # Connect to imap server and start tls
    imapserver = imaplib.IMAP4(host=secrets.townIMAPServer,port=143)
    imapserver.starttls()
    print("Connecting...")
    try:
        # Login via secrets credentials
        imapserver.login(secrets.opendatausername,secrets.opendatapassword)
        print("Connected.")
    except:
        print("Not connected.")
    # Return server
    return imapserver

# Connect to imap server and start tls
try:
    mailBox = login_and_write()
    mailBox.select(mailbox='Inbox/Self-Check', readonly=False)
    # Only 'Unseen' emails - may not be working as expected
    messages = mailBox.search(None, 'UNSEEN')
    result, data = mailBox.uid('search', None, searchQuery)
    ids = data[0]
    # list of uids
    id_list = ids.split()
    i = len(id_list)
    for x in range(i):
        latest_email_uid = id_list[x]

        # What are the flags now?
        #typ, response = mailBox.fetch(latest_email_uid, '(FLAGS)')
        #print(f'Flags before: {response}')
        # fetch the email body (RFC822) for the given ID
        result, email_data = mailBox.uid('fetch', latest_email_uid, '(RFC822)')
        raw_email = email_data[0][1]

        # converts byte literal to string removing b''
        raw_email_string = raw_email.decode('utf-8')
        email_message = email.message_from_string(raw_email_string)

        # downloading attachments
        for part in email_message.walk():
            # this part comes from the snipped I don't understand yet...
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue
            fileName = part.get_filename()

            if bool(fileName):
                filePath = output_folder.joinpath(fileName)
                if not os.path.isfile(filePath) :
                    fp = open(filePath, 'wb')
                    fp.write(part.get_payload(decode=True))
                    fp.close()
                    print(f'{filePath} saved')
                    mailBox.store(latest_email_uid, '+FLAGS', '\\Deleted')
                    # What are the flags now?
                    # typ, response = mailBox.fetch(latest_email_uid, '(FLAGS)')
                    # print(f'Flags after: {response}')
    mailBox.close()
    mailBox.logout()
except:
    print("email retrieval failed")
    exit(1)