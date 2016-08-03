#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      sdurgawad
#
# Created:     01/08/2016
# Copyright:   (c) sdurgawad 2016
# Licence:     <your licence>
#-------------------------------------------------------------------------------

import email
import imaplib
import os
import shutil
import sys
from datetime import datetime


def open_connection(verbose=False):

    connectionParams = {}
    f = open('kitkat', 'r') # Open the files to read connection parameters
    for line in f:
        parameters = line.strip().split('=') # split around the = sign
        if len(parameters) > 1: # we have the = sign in there
            connectionParams[parameters[0]] = parameters[1]

    # Connect to the gmail server
    hostname = connectionParams['gmailServer']
    if verbose: print 'Connecting to', hostname
    connection = imaplib.IMAP4_SSL(hostname)

    # Login to our account
    username = connectionParams['username']
    password = connectionParams['password']
    if verbose: print 'Logging in as', username
    connection.login(username, password)

    downloadDirectory = connectionParams['downloadDirectory']

    return connection, downloadDirectory


def downloadAttachments(imapSession, directoryPath, criteria, criteriaValue):

    imapSession.select('inbox')
    # resp, items = imapSession.search(None, 'FROM', '"sdurgawad@altimetrik.com"')
    # resp, items = imapSession.search(None, 'SUBJECT', '"Your airtel Bill for Mobile Number: 8056110703"')

    #date = datetime.strptime("01/07/2015", "%d/%m/%Y").strftime("%d-%b-%Y")
    #resp, items = imapSession.search(None, 'SENTSINCE', '{date}'.format(date=date))
    #resp, items = imapSession.search(None, 'SENTON', '{date}'.format(date=date))

    resp, items = imapSession.search(None, criteria, criteriaValue)

    items = items[0].split() # getting the mails id

    counter = 0

    for emailid in items:

        # fetching the mail, "`(RFC822)`" means "get the whole stuff", but you can ask for headers only, etc
        resp, data = imapSession.fetch(emailid, "(RFC822)")

        # getting the mail content
        email_body = data[0][1]

        # parsing the mail content to get a mail object
        mail = email.message_from_string(email_body)

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

            att_path = os.path.join(directoryPath, filename)

            if os.path.isfile(att_path):
                filename = str(counter) + filename
                att_path = os.path.join(directoryPath, filename)
                counter += 1


            # Check if its already there
            if not os.path.isfile(att_path) :
                # finally write the stuff
                fp = open(att_path, 'wb')
                fp.write(part.get_payload(decode=True))
                fp.close()


def exitApp():
    """ Gets exit from the program """
    print "\n Successfully exited the program"
    exit


options = {
            1 : downloadAttachments,
            2 : downloadAttachments,
            3 : downloadAttachments,
            4 : downloadAttachments,
            5 : downloadAttachments,
            6 : exitApp,
        }


def PrintOptions():
    """ This function lists the options """
    print("""
            GET THE ATTACHMENTS BASED ON BELOW CRITERIA:

            1. From a Mail ID
            2. From a SUBJECT
            3. on or after this Date
            4. on or before this Date
            5. On this Date
            6. Between these two entered DATES
            7. Exit  """)


def main():

    serverConnection, downloadDirectory = open_connection(True)

    while True:
        PrintOptions()
        option = int(input("Enter the option from the list: "))

        if options.has_key(option):
            # Call zipfolder function
            if option == 1:
                emailID = str(raw_input("Enter the valid email ID"))
                options[option](serverConnection, downloadDirectory, 'FROM', emailID)

            # Call extractArchiveFile function
            elif option == 2:
                subjectString = str(raw_input("Enter the Subject string"))
                options[option](serverConnection, downloadDirectory, 'SUBJECT', subjectString)

            # Call listFilesFromArchive function
            elif option == 3:
                dateSentSince = str(input("Enter the Date, to get the attachments on or after that date in the format %d/%m/%Y"))
                options[option](serverConnection, downloadDirectory, 'SENTSINCE', dateSentSince)

            # Call deleteFilesFromArchive function
            elif option == 4:
                dateSentBefore = str(input("Enter the Date, to get the attachments on or before that date, in the format %d/%m/%Y"))
                options[option](serverConnection, downloadDirectory, 'SENTBEFORE', dateSentBefore)

            # Call deleteFilesFromArchive function
            elif option == 5:
                dateSentOn = str(input("Enter the Date, to get the attachments on that date, in the format %d/%m/%Y"))
                date = datetime.strptime(str(dateSentOn), "%d/%m/%Y").strftime("%d-%b-%Y")
                resp, items = imapSession.search(None, 'SENTON', '{date}'.format(date=date))
                options[option](serverConnection, downloadDirectory, resp, items)

            # Call exitApp function
            elif option == 6 :
                options[option]()
                break

        continue

    serverConnection.close()
    serverConnection.logout()


if __name__ == '__main__':
    main()


