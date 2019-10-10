# Dependencies
import imaplib
import email
import secrets, filename_secrets
import PDC_lookup
import pathlib
import os, csv, re, datetime, traceback, datetime

# Search Parameters
OPS_code = re.compile(r'OPS\d')
CAD_code = re.compile(r'CAD:\d{6,10}')
PDC_code = re.compile(r'\d{1,3}[A-Z]{1}\d{1,2}')
extraneous_chars = re.compile(r"(b'|=|\\r|\\n|'|\"|\[|\]|\\|\r\n|CAD:)")
ems_match = re.compile(r"(SICK|INJU|FAINT|UNCON|PAIN|HEART|OVERDOSE)")

# Function to login to IMAP
def IMAP_login():
    # Connect to imap server and start tls
    #imapserver = imaplib.IMAP4(host=secrets.townIMAPServer,port=143)
    imapserver = imaplib.IMAP4(host=secrets.townIMAPServer2,port=143)
    imapserver.starttls()
    print("Connecting...")
    try:
        # Login via secrets credentials
        imapserver.login(secrets.odsuser,secrets.odspass)
        print("Connected.")
    except:
        raise NameError('IMAP connection Failed')
    return imapserver

# Function to match PDC code
def PDC_mapper(source, textDescription):
    global PDC_code
    try:
        if re.match(PDC_code, source):
            PDC_match = re.split(r'[A-Z]',  source, maxsplit=1)
            if int(PDC_match[0]) > 0 and int(PDC_match[0]) < 50:
                textDescription = "EMS"
            elif int(PDC_match[0]) > 50 and int(PDC_match[0]) < 100:
                textDescription = PDC_lookup.PDC[int(PDC_match[0])]
            elif int(PDC_match[0]) > 100:
                textDescription = "POLICE"
    except (IndexError, KeyError):
        textDescription = "UNKNOWN/UNSPECIFIED"
    return textDescription

# Function to fetch data from chosen mailbox folder
def etl_data(server):
    global OPS_code, PDC_code
    # Create Path to staging directory
    stagingPath = pathlib.Path(filename_secrets.productionStaging)

    # Open csv
    outputFilename = stagingPath.joinpath('fire_dept_dispatches.csv')
    outputFile = open(outputFilename, 'a+')
    csv_header = ["CAD","Address","City","Type of Incident","Date","PDC"]
    info_sheet = csv.writer(outputFile)

    # If there's no header, write headers
    if os.stat(outputFilename).st_size == 0:
        info_sheet.writerow(csv_header)
    # Selects inbox as target
    server.select(mailbox='Inbox')

    # Select emails since yesterday
    yesterday = datetime.date.today() - datetime.timedelta(days=1)
    searchQuery = '(FROM "Orange Co EMS Dispatch" SINCE "' + yesterday.strftime('%d-%b-%Y') + '")'
    # test for all emails below
    #searchQuery = '(FROM "Orange Co EMS Dispatch")'
    result, data = server.uid('search', None, searchQuery)
    message_list = data[0].split()
    i = len(message_list)
    print(f'{i} new messages found')

    # Fetch Envelope data which contains date received
    for x in range(i):
        latest_email_uid = message_list[x]
        result, email_data = server.uid('fetch', latest_email_uid, '(RFC822)')
        raw_email = email_data[0][1]

       # converts byte literal to string removing b''
        raw_email_string = raw_email.decode('utf-8')
        email_message = email.message_from_string(raw_email_string)
        #print(email_message)
        email_date = datetime.datetime.strptime(email_message._headers[7][1][5:], '%d %b %Y %X %z')

        # output list
        output_fields = [None]*6
        # fetch text content
        split_text = email_message._payload.split(";")
        # Input format is expected to be
        # CAD;Address;City;Text Desc;OPS;PDC
        if not re.match(CAD_code, split_text[0]):
            # Record is invalid if not started with a CAD code
            print(f'Incomplete record: {split_text}')
            continue
        if len(split_text) < 4:
            # if we don't have complete CAD record, skip to next email
            print(f'Incomplete record: {split_text}')
            continue
        try:
            output_fields[0] = re.sub(extraneous_chars, '', split_text[0])
            output_fields[1] = re.sub(extraneous_chars, '', split_text[1])
            output_fields[2] = re.sub(extraneous_chars, '', split_text[2])

            try:
                textDescription = re.sub(extraneous_chars, '', split_text[3])
                if len(split_text) < 5:
                    # no codes
                    if re.match(ems_match, split_text[3]):
                        textDescription = "EMS"
                    else:
                        textDescription = re.sub(extraneous_chars, '', split_text[3])
                if len(split_text) > 4:
                    # OPS code
                    split_text[4]    = re.sub(extraneous_chars, '', split_text[4])
                    if re.match(OPS_code, split_text[4]):
                        if split_text[4] == "OPS1":
                            textDescription = "EMS"
                        elif split_text[4] == "OPS2":
                            textDescription = "FIRE"
                    elif re.match(PDC_code, split_text[4]):
                        textDescription = PDC_mapper(split_text[4], textDescription)

                if len(split_text) > 5:
                    # PDC 
                    split_text[5] = re.sub(extraneous_chars, '', split_text[5])
                    textDescription = PDC_mapper(split_text[5], textDescription)        

            except (IndexError, KeyError):
                textDescription = "UNKNOWN/UNSPECIFIED"
                
            output_fields[3] = textDescription.upper()
            output_fields[4] = datetime.datetime.strftime(email_date, '%Y-%m-%d %H:%M')            
            #output_string = str(output_fields)
            # Clean string up and write data to csv
            print(output_fields)
            info_sheet.writerow(output_fields)
        except IndexError:
            pass

    # Close CSV
    outputFile.close()
    server.logout()

# Main
exchange_mail = IMAP_login()
etl_data(exchange_mail)