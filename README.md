# OD-email-reader

[![Codacy Badge](https://api.codacy.com/project/badge/Grade/5d2c2cb5fd494ac492ae2a4b6dd5fc6e)](https://app.codacy.com/app/TownofChapelHill/OD-email-reader?utm_source=github.com&utm_medium=referral&utm_content=townofchapelhill/OD-email-reader&utm_campaign=Badge_Grade_Dashboard)

## Designed to read the contents of the Open Data email inbox.
## Contains specific code for processing Fire Dispatch content

### Goal 
Create scripts to process IMAP email from exchange

### Purpose 
Import email contents into CSV files

Save attachments to a filesystem folder

Deduplicate CSV files

### Methodology 
Uses imapclient and email modules

Powershell script

### Transformations
#### csv_file_dedup.ps1
sorts and dedups a csv file. Input, Output, and sort field specified as parameters

#### read-email.py
reads emails from INBOX of odsuser

creates .../fire_dept_raw_dispatches.csv by parsing emails

creates .../fire_dept_dispatches_clean.csv as a cleanup phase

#### save-email-attachment.py
reads emails from INBOX/Self-Check of the opendata user

save attachments to .../Selfchecks which are then processed for the portal (https://github.com/townofchapelhill/self-check)

### Constraints
Requires imapclient via ```pip install imapclient```

Requires imaplib via ```pip install imaplib```
