# Dependencies
import imaplib
import email
import secrets, filename_secrets
# import PDC_lookup
import pathlib
import os, csv, re, datetime, traceback, datetime, ssl
import pandas as pd
from pandas.errors import ParserError
import logging

# Search Parameters
OPS_code = re.compile(r'OPS\d')
PDC_code = re.compile(r'\d{1,3}[A-Z]{1}\d{1,2}')
extraneous_chars = re.compile(r"(b'|=|\\r|\\n|'|\"|\[|\]|\\|\r\n|CAD:)")

# Function to login to IMAP
def IMAP_login():

    # Connect to imap server and start tls
    imapserver = imaplib.IMAP4(host=secrets.townIMAPServer,port=143)
    imapserver.starttls()
    print("Connecting...")
    try:
        # Login via secrets credentials
        imapserver.login(secrets.odsuser,secrets.odspass)
        print("Connected.")
    except:
        raise NameError('IMAP connection Failed')
    return imapserver


# Function to fetch data from chosen mailbox folder
def etl_data(server):
    global OPS_code, PDC_code

    # Create Path to staging directory
    stagingPath = pathlib.Path(filename_secrets.productionStaging)
    # Open csv
    outputFilename = stagingPath.joinpath('fire_dept_raw_dispatches_test.csv')
    info_sheet = open(outputFilename, 'w')
    # Selects inbox as target
    server.select(mailbox='Inbox')
    # Select emails since yesterday
    yesterday = datetime.date.today() - datetime.timedelta(days=7)
    searchQuery = '(FROM "Orange Co EMS Dispatch" SINCE "' + yesterday.strftime('%d-%b-%Y') + '")'
    # If there's no header, write headers
    if os.stat(outputFilename).st_size == 0:
        info_sheet.write("CAD,City,Type of Incident\n")

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
        output_fields = [None]*5
        # fetch text content
        split_text = email_message._payload.split(";")
        # Input format is expected to be
        # CAD;Address;City;Text Desc;OPS;PDC

        output_fields[0] = re.sub(extraneous_chars, '', split_text[0])
        output_fields[1] = re.sub(extraneous_chars, '', split_text[1])
        output_fields[2] = re.sub(extraneous_chars, '', split_text[2])
        textDescription = "UNKNOWN/UNSPECIFIED"
        try:
            if re.match(PDC_code, split_text[5]):
                PDC_match = re.split(r'[A-Z]',  split_text[5], maxsplit=1)
                if int(PDC_match[0]) > 0 and int(PDC_match[0]) < 50:
                    textDescription = "EMS"
                elif int(PDC_match[0]) > 50 and int(PDC_match[0]) < 100:
                    textDescription = "FIRE"
                    elif int(PDC_match[0]) > 100:
                        textDescription = "Police"
                #textDescription = PDC_lookup.PDC[int(PDC_match[0])]
            elif re.match(OPS_code, split_text[5]):
                if split_text[5] is "OPS1":
                    textDescription = "EMS"
                elif split_text[5] is "OPS2":
                    textDescription = "FIRE"            
        except (IndexError, KeyError):
            textDescription = re.sub(extraneous_chars, '', split_text[3])
        output_fields[3] = textDescription.upper()
        output_fields[4] = datetime.datetime.strftime(email_date, '%Y-%m-%d %H:%M')

        # Convert item to string
        output_string = str(output_fields)
        # Clean string up and write data to csv
        print(output_string)
        info_sheet.write(output_string + "\n")
    # Close CSV being used
    info_sheet.close()
    # Call cleanup_csv function using dates list
    #cleanup_csv(datetimes)
    # Call logout function
    server.logout()

# Function to clean up CSV that is created
def cleanup_csv(dateslist):
    print("Cleaning CSV...")
    # Create Path to staging directory
    stagingPath = pathlib.Path(filename_secrets.productionStaging)
    csvFilename = stagingPath.joinpath('fire_dept_raw_dispatches.csv')
    # Create pandas dataframe from original csv
    df = pd.read_csv(csvFilename)
    # , error_bad_lines=False, warn_bad_lines=True

    # Delete PII rows and drop duplicate records
    del df["ID"]
    del df["ID2"]
    df['Dates'] = pd.Series(dateslist, index = df.index[:len(dateslist)])
    df.drop_duplicates(keep='first')
    print(df)
    outputFilename = stagingPath.joinpath('fire_dept_dispatches_clean.csv')
    clean_file = open(outputFilename, "a+")
    # Write headers to blank clean file
    if os.stat(outputFilename).st_size == 0:
        clean_file.write(",CAD,Address,City,Type of Incident,Dates\n")
    # df.drop("Address", axis=1, inplace=True)

    # Write dataframe to new, finalized csv
    df.to_csv(clean_file, mode='w', header=False)
    print("CSV cleaned and rewritten.")


# Main function to call other functions
def main():

    # Open log file
    #logfilePath = pathlib.Path(filename_secrets.logfilesDirectory)
    #fireLog = logfilePath.joinpath('fire_dispatch_log.txt')
    #fire_log = open(fireLog, "a+")
    # try:
        # Set var to hold what is returned from long_and_write()
    exchange_mail = IMAP_login()
        # Call etl_data() using the exchange imap server
    etl_data(exchange_mail)
        # Log failures and successes
    #fire_log.write(str(datetime.datetime.now()) + "\n" + "Logged into Exchange IMAP and saved emails." + "\n")
    # except:
        # fire_log.write(str(datetime.datetime.now()) + "\n" + "There was an issue when trying to log in and save emails." + "\n" + traceback.format_exc() + "\n")

# Call main
#main()
exchange_mail = IMAP_login()
etl_data(exchange_mail)