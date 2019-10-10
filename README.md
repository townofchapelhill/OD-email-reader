# OD-email-reader

[![Codacy Badge](https://api.codacy.com/project/badge/Grade/5d2c2cb5fd494ac492ae2a4b6dd5fc6e)](https://app.codacy.com/app/TownofChapelHill/OD-email-reader?utm_source=github.com&utm_medium=referral&utm_content=townofchapelhill/OD-email-reader&utm_campaign=Badge_Grade_Dashboard)

## Designed to read the contents of the Open Data email inbox.
## Contains specific code for selecting emails for content areas

### Goal 
Create scripts to process IMAP email from exchange

### Purpose 
1. Import email contents into CSV files for downstream reuse
2. Save attachments to a filesystem folder

Deduplicate CSV files

### Methodology 
Uses imapclient, imaplib and email modules

Powershell script

### Transformations
#### csv_file_dedup.ps1
sorts and dedups a csv file. Input, Output, and sort field specified as parameters

#### IMAP_connection_test.py
test connectivity to town IMAP services. Dump a trace on connection failure.

#### read-email.py (deprecated)
#### read-email-ems.py
reads emails from INBOX of odsuser
retrieve EMS records forwarded via twitter 

creates .../fire_dept_dispatches.csv by parsing emails

#### save-email-attachment.py
reads emails from INBOX/Self-Check of the opendata user

save attachments to .../Selfchecks which are then processed for the portal (https://github.com/townofchapelhill/self-check)

### Constraints
Requires imapclient via ```pip install imapclient```

Requires imaplib via ```pip install imaplib```
