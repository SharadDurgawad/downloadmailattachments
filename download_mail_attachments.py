#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      sdurgawad
#
# Created:     29/07/2016
# Copyright:   (c) sdurgawad 2016
# Licence:     <your licence>
#-------------------------------------------------------------------------------

import email, getpass, imaplib, os, shutil, sys

user = raw_input("Enter your GMail username:")
pwd = getpass.getpass("Enter your password:")


# connecting to the gmail imap server
imapSession = imaplib.IMAP4_SSL("imap.gmail.com")
imapSession.login(user,pwd)

# Select a label you wish to download the attachments from
imapSession.select("Altimetrik") # here you a can choose a mail box like INBOX instead

# use imapSession.list() to get all the mailboxes
# imapSession.list()

resp, items = imapSession.search(None, 'FROM', '"sdurgawad@altimetrik.com"') # you could filter using the IMAP rules here (check http://www.example-code.com/csharp/imap-search-critera.asp)
items = items[0].split() # getting the mails id

counter = 0

# Directory path to download the attachments
download_dir = raw_input("Enter your valid directory path:")

# If the entered directory path does not exists then create the same
if not os.path.exists(download_dir):
    os.makedirs(download_dir)

for emailid in items:
    resp, data = imapSession.fetch(emailid, "(RFC822)") # fetching the mail, "`(RFC822)`" means "get the whole stuff", but you can ask for headers only, etc
    email_body = data[0][1] # getting the mail content
    mail = email.message_from_string(email_body) # parsing the mail content to get a mail object

    # Check if any attachments at all
    if mail.get_content_maintype() != 'multipart':
        continue

    print "["+mail["From"]+"] :" + mail["Subject"]

    # we use walk to create a generator so we can iterate on the parts and forget about the recursive headach
    for part in mail.walk():
        # multipart are just containers, so we skip them
        if part.get_content_maintype() == 'multipart':
            continue

        # is this part an attachment ?
        if part.get('Content-Disposition') is None:
            continue

        filename = part.get_filename()
        #counter = 1
        """
        # if there is no filename, we create one with a counter to avoid duplicates
        if not filename:
            filename = 'part-%03d%s' % (counter, 'bin')
            counter += 1
        """

        att_path = os.path.join(download_dir, filename)

        if os.path.isfile(att_path):
            filename = str(counter) + filename
            att_path = os.path.join(download_dir, filename)
            counter += 1


        # Check if its already there
        if not os.path.isfile(att_path) :
            # finally write the stuff
            fp = open(att_path, 'wb')
            fp.write(part.get_payload(decode=True))
            fp.close()