# Dependencies
import imaplib
import email
import secrets, filename_secrets
import pathlib
import os, csv, re, datetime, traceback, datetime, ssl
import pandas as pd
from pandas.errors import ParserError
import logging


# Function to login to IMAP
def login_and_write():

    # Connect to imap server and start tls
    imapserver = imaplib.IMAP4(host=secrets.townIMAPServer,port=143)
    imapserver.starttls()
    print("Connecting...")
    try:
        # Login via secrets credentials
        imapserver.login(secrets.odsuser,secrets.odspass)
        print("Connected.")
    except:
        print("Not connected.")
    # Return server
    return imapserver


# Function to fetch data from chosen mailbox folder
def etl_data(server):

    # Create list to later hold received datetimes for messages
    datetimes = []
    # Create list to hold Cities
    cities_list = ["CHAPEL HILL", "DURHAM", "MEB", "HILLSBOROUGH", "CARRBORO", "UNKN"]
    # Create Path to staging directory
    stagingPath = pathlib.Path(filename_secrets.productionStaging)
    # Open csv
    outputFilename = stagingPath.joinpath('fire_dept_raw_dispatches.csv')
    info_sheet = open(outputFilename, 'w')
    # Selects inbox as target
    server.select(mailbox='Inbox')
    # Select emails since yesterday
    yesterday = datetime.date.today() - datetime.timedelta(days=1)
    searchQuery = '(FROM "Orange Co EMS Dispatch" SINCE "' + yesterday.strftime('%d-%b-%Y') + '")'
    # If there's no header, write headers
    if os.stat(outputFilename).st_size == 0:
        info_sheet.write("CAD,Address,City,Type of Incident,ID,ID2\n")

    result, data = server.uid('search', None, searchQuery)
    message_list = data[0].split()
    i = len(message_list)
    print(f'{i} new messages: {message_list}')

    # Fetch Envelope data which contains date received
    for x in range(i):
        latest_email_uid = message_list[x]
        result, email_data = server.uid('fetch', latest_email_uid, '(RFC822)')
        raw_email = email_data[0][1]

        # converts byte literal to string removing b''
        raw_email_string = raw_email.decode('utf-8')
        email_message = email.message_from_string(raw_email_string)
        print(email_message)
        email_date = datetime.datetime.strptime(email_message._headers[7][1][5:], '%d %b %Y %X %z')
        # If item in list is a datetime object
        if isinstance(email_date, datetime.datetime):
            # Add dates to a list for later use
            datetimes.append(email_date)

        # fetch text content
        split_text = email_message._payload.split(";")
        print(split_text)

        # Check for city column errors
        try:
            if split_text[2] not in cities_list:
                # If item is not a city, insert 'UNKN, '
                split_text[2] = "UNKN, \n" + str(split_text[2])
        except:
            pass
        # Convert item to string
        split_text = str(split_text)
        # Clean string up and write data to csv
        split_text = split_text.replace("CAD:","").replace("b'","").replace("=","").replace('\\r', '').replace('\\n', '').replace("'","").replace('"', "").replace('[','').replace(']','').replace("\\", "")
        print(split_text)
        info_sheet.write(split_text + "\n")
    # Close CSV being used
    info_sheet.close()
    # Call cleanup_csv function using dates list
    cleanup_csv(datetimes)
    # Call logout function
    logout(server)


# Logs out of remote session
def logout(server):

    server.logout()
    b'Logging out'


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
    logfilePath = pathlib.Path(filename_secrets.logfilesDirectory)
    fireLog = logfilePath.joinpath('fire_dispatch_log.txt')
    fire_log = open(fireLog, "a+")
    # try:
        # Set var to hold what is returned from long_and_write()
    exchange_mail = login_and_write()
        # Call etl_data() using the exchange imap server
    etl_data(exchange_mail)
        # Log failures and successes
    fire_log.write(str(datetime.datetime.now()) + "\n" + "Logged into Exchange IMAP and saved emails." + "\n")
    # except:
        # fire_log.write(str(datetime.datetime.now()) + "\n" + "There was an issue when trying to log in and save emails." + "\n" + traceback.format_exc() + "\n")

# Call main
main()
