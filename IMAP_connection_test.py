''' IMAP CONNECCTION TEST
    attempts to connect to am IMAP resource
    on failure - print out the connection trace
'''
# Dependencies
import imaplib
import email
import secrets, filename_secrets
import os, re, datetime, traceback

# Function to login to IMAP
def IMAP_login(emailServer):

    # Connect to imap server and start tls
    imapserver = imaplib.IMAP4(host=emailServer,port=143)
    imapserver.starttls()
    print("Connecting...")
    try:
        # Login via secrets credentials
        imapserver.login(secrets.odsuser,secrets.odspass)
        print("Connected.")
    except:
        raise NameError('IMAP connection Failed')
        traceback.print_exc()
    return imapserver

# Main
print(f' Connecting to {secrets.townIMAPServer}')
exchange_mail = IMAP_login(secrets.townIMAPServer)
exchange_mail.logout()

print(f' Connecting to {secrets.townIMAPServer2}')
exchange_mail = IMAP_login(secrets.townIMAPServer2)
exchange_mail.logout()