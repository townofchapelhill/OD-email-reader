# Read emails from the subfolder "purchase-suggestion" and parse for known fields

# Dependencies
import imaplib
import email
import secrets, filename_secrets
import os, csv, re, datetime, traceback, datetime, ssl, pprint, pathlib
from html.parser import HTMLParser

# Globals
result_dict = {}
end_colon = re.compile(':$')
i = 0
mapping_keys = {'14 digit library card number or CHCCS ID:': 'Card Number',
            'Audience:': 'Audience',
            'Email:': 'Email',
            'Format:': 'Format',
            'First Name:': 'FirstName',
            'Last Name:': 'Last Name',
            'Item author or performer:': 'author',
            'Item title :': 'Title'}
csv_headers = ['Card Number', 'Audience', 'FirstName', 'LastName', 'Email', 'Author', 'Title']
csv_line = {}

# Class to help parse and format HTML elements
class myHTMLParser(HTMLParser):
    def __init__(self):
        # initialize the base class
        HTMLParser.__init__(self)
        self.form_data = []
        self.message_count = 0

    def delete(self, i):
        if i > len(form_data):
            raise IndexError
        else:
            del self.form_data[i]
            self.message_count -= 1

    #def handle_starttag(self, tag, attrs):
        #print(f'Encountered a start tag: {tag}')

    #def handle_endtag(self, tag):
        #print(f'Encountered a end tag: {tag}')

    def handle_data(self, data):
        parsed_data = data.replace('\s*','').replace('\r','').replace('\n','')
        if (len(parsed_data) > 2):
            self.form_data.append(parsed_data)
            self.message_count += 1

# Main function to call other functions
def main():

    # try:
        # Connect to imap server
    exchange_mail = login_imap()
        # Call etl_data() using the exchange imap server
    etl_data(exchange_mail)

# Function to login to IMAP
def login_imap():

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


# Function to fetch data from chosen mailbox folder
def etl_data(server):

    # Create Path to staging directory
    stagingPath = pathlib.Path(filename_secrets.productionStaging)
    outputFilename = stagingPath.joinpath('purchase-suggestions.csv')
    # Open output csv
    info_sheet = open(outputFilename, 'a+')
    # If there's no header, write headers
    if os.stat(outputFilename).st_size == 0:
        info_sheet.write(csv_headers)

    # Selects box as target
    result, message_count = server.select(mailbox='Inbox/purchase-suggestion', readonly=False)
    # Select emails since yesterday
    yesterday = datetime.date.today() - datetime.timedelta(days=1)
    searchQuery = '(SINCE "' + yesterday.strftime('%d-%b-%Y') + '")'
    result, data = server.uid('search', None, searchQuery)
    message_list = data[0].split()
    i = len(message_list)
    print(f'There are {i} new messages')

    # Fetch Envelope data
    for x in range(i):
        latest_email_uid = message_list[x]
        result, email_data = server.uid('fetch', latest_email_uid, '(RFC822)')
        raw_email = email_data[0][1]

        # converts byte literal to string removing b''
        raw_email_string = raw_email.decode('utf-8')
        email_message = email.message_from_string(raw_email_string)
        parser = myHTMLParser()
        parser.feed(email_message._payload)
        i = 0
        type(parser.form_data)
        while True:
            # build a dictionary of the result set
            try:
                if end_colon.search(parser.form_data[i]) is not None:
                    result_dict[parser.form_data[i]] = parser.form_data[i+1]
            except IndexError:
                break
            i += 1

        for key in mapping_keys.keys():
            csv_line[mapping_keys[key]] = result_dict[key]
        
        pprint.pprint(csv_line)

        info_sheet.write(csv_line.values() + '\n')
    # Close CSV being used
    # Call cleanup_csv function using dates list
    #cleanup_csv(datetimes)
    # Call logout function
    server.logout()


# Function to clean up CSV that is created
def cleanup_csv(dateslist):
    print("Cleaning CSV...")
    # Create Path to staging directory
    stagingPath = pathlib.Path(filename_secrets.productionStaging)
    csvFilename = stagingPath.joinpath('purchase-suggestions.csv')
    # Create pandas dataframe from original csv
    df = pd.read_csv(csvFilename)
    # , error_bad_lines=False, warn_bad_lines=True

    # Delete PII rows and drop duplicate records
    del df["ID"]
    del df["ID2"]
    df['Dates'] = pd.Series(dateslist, index = df.index[:len(dateslist)])
    df.drop_duplicates(keep='first')
    print(df)

    outputFilename = stagingPath.joinpath('purchase-suggestions.csv')
    clean_file = open(outputFilename, "a")
    # Write headers to blank clean file
    if os.stat(outputFilename).st_size == 0:
        clean_file.write(",CAD,Address,City,Type of Incident,Dates\n")
    # df.drop("Address", axis=1, inplace=True)

    # Write dataframe to new, finalized csv
    df.to_csv(clean_file, mode='w', header=False)
    print("CSV cleaned and rewritten.")


# Call main
main()
