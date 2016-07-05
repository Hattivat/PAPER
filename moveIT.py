# -*- coding: utf-8 -*-
"""
Created on Fri Feb 12 12:27:02 2016
Revised on Mon Mar 07 15:10:00 2016

@author: Jan Janiczek
"""
# import os
import win32com.client as client
# import string

# SENDER = "Jan Janiczek\nIT Trainee\nPrivileged Access Provisioning"
# username = input("your username: \n")
# password = input("your password: \n")

SHAREDMAILBOX = "xxxxxxxxxxxx"
SIGNATURE = "Jan Janiczek\nIT Trainee\nPrivileged Access Provisioning\nxxxxx \
xxxxxxx\nxxxxx\nxxxxxxxxxxxxxxxxxx, xxxxxxxxxxxxx, xxxxxx\nEmail \
xxxxxxxxxxxxxx@xxxxxxxxx"


def parse_text(listmail):
    """Reads an email and extracts the required information from it - worker,
    manager, their MUD IDs, PAMS groups the worker belongs to.

    takes:
    listmail - a list containing a MoveIT email request split on newlines;
    it is assumed that checking whether that is the case will be handled
    outside of this function prior to calling it.

    returns:
    worker: a string,
    worker_mud: a string,
    manager_f_name: a string,
    manager_mud: a string,
    accesses: a list of strings"""

    access_time = False  # flag to signal we are inside the list of transferee accesses
    formatting_fail = False  # flag to signal that new manager's MUD is on a newline

    for line in listmail:

        if formatting_fail:  # to handle cases where MUD is on a newline
            new_manager_mud = line[1:-4]
            formatting_fail = False

        if line.startswith("MoveIT Transferee"):  # line with that beginning contains transferee's ID
            parenindex = line.find("(")
            parenindex2 = line.find(")")
            if parenindex > 18 and parenindex2 > parenindex:  # sanity check
                worker = line[18:parenindex-1]
                worker_mud = line[parenindex+1:parenindex2]
                worker_temp = worker.split()
                worker_f_name = worker_temp[0]
            else:
                # never seen it happen, but doesn't hurt to handle the possibility
                raise TypeError("No ID for the transferee?!")

        if access_time and line == "\r":  # empty line means the end of accesses list
            access_time = False

        elif access_time:  # as long as no newline appears, read accesses
            linecontent = line.split("---")
            # The line below will only ever be executed after the lines below
            # it trigger on a previous loop and assign an empty list to accesses,
            # so the IDE's silly warning can be safely disregarded.
            accesses.append(linecontent[1].strip())

        if line.startswith("account name"):  # the line before the list of accesses is always like this
            access_time = True
            accesses = []

        if line.startswith("The current manager"):  # Extract current manager's ID (always provided)
            cur_name_begin = line.find("is") + 2
            cur_name_end = line.find("(") - 1
            cur_manager_name = line[cur_name_begin:cur_name_end]
            cur_manager_name = cur_manager_name.split()
            cur_manager_f_name = cur_manager_name[0]
            cur_manager_mud = line[cur_name_end+2:-2]

        if line.startswith("The new manager"):  # Extract new manager's ID (if there is one)
            new_name_begin = line.find(":") + 2
            new_name_end = line.find("(", 78, -1) - 1
            new_manager_name = line[new_name_begin:new_name_end]
            try:  # if there is a new manager, extract the data
                new_manager_name_split = new_manager_name.split()
                new_manager_f_name = new_manager_name_split[0]
                new_manager_mud = line[new_name_end+2:-3]
            except IndexError:  # if there isn't one, [0] index will result in IndexError
                new_manager_f_name = ""  # substitute with empty strings for easy detection
                new_manager_mud = ""
            if new_name_end == -2:  # Indicates that no MUD ID is on the current line
                formatting_fail = True

    else:  # Check which of the managers needs to be contacted
        if new_manager_f_name != "":  # There is a new manager, so he is to be contacted
            manager_f_name = new_manager_f_name
            manager_mud = new_manager_mud
        elif cur_manager_f_name != "":  # No new manager, so the current (previous) one has stayed
            manager_f_name = cur_manager_f_name
            manager_mud = cur_manager_mud
        else:  # Hopefully will not happen again, but has occurred once in my experience
            raise TypeError("No manager!!!!")
    # returns four strings and one list of string
    return worker, worker_f_name, worker_mud, manager_f_name, manager_mud, accesses


def create_mail(inputtext):
    """Takes an unformatted string of text from an email, splits it on newlines
    to enable line-by-line reading, and then parses it using the parse_text
    function.

    takes:
    inputtext: a string

    returns:
    result: a string
    """

    splitmail = inputtext.split("\n")
    worker, worker_f_name, worker_mud, manager_f_name, manager_mud, \
    accesses = parse_text(splitmail)

    subject = "Confirmation of privileged access - {0}".format(worker)
    accesses_list = "\n".join(x for x in accesses)
    mailtext = "Dear {0},\n\nWe are contacting you as you are listed as the new \
manager for {1}, who recently migrated and/or moved (Location and/or Position).\
\nDue to the Sarbanes Oxley Compliance audit requirements we need an email \
confirmation that {2} should either retain or not keep the following privileged \
access:\n\n{3}\n\nIf we do not get a response from you within 3 working days \
the privileged access will be terminated.\nPlease direct your replies to the \
xxxxxxxxxxxxxxxxxxxxxxxxx mailbox.\n\nThank you in advance for your \
cooperation.\n\nKind Regards,".format(manager_f_name, worker,
    worker_f_name, accesses_list)

    return manager_mud, worker_mud, subject, mailtext


def process_mail(message, outlook, sharedmail, signature):
    """Takes a message, an outlook application, and the name of shared mail to
    be used, and then creates an appropriately named output file, and sends an 
    email resulting from processing that message with the create_mail function 
    to the appropriate recipients. Briefly opens a new message window in Outlook,
    as this is, despite being ugly, apprently the actual best solution to extract
    the text of the user's signature without knowing under what filename it is
    saved (this is user-defined, so it can, and does, vary).

    takes:
    message: an outlook email object
    outlook: an outlook application object
    sharedmail: a string containing the name of shared account to use as
    the on-behalf-of email
    signature: the signature appropriate for the sender

    returns:
    nothing; saves email templates as files and prints the filenames to the
    console"""

    manager, transferee, subject, text = create_mail(message.body)
    signedtext = text + "\n" + signature
    olMailItem = 0x0
    newMail = outlook.CreateItem(olMailItem)
    newMail.Subject = subject
    newMail.Body = signedtext
    newMail.To = manager
    newMail.CC = transferee
    newMail.SentOnBehalfOfName = sharedmail
    newMail.Send()
    print("Email sent to {0}".format(manager))

def main():
    # A flag for whether any MoveIT requests have been found and processed
    moveitsent = False
    # Set up the connection with Outlook
    outlook = client.Dispatch("Outlook.Application")
    outlookMAPI = outlook.GetNamespace("MAPI")
    recipient = outlookMAPI.CreateRecipient(SHAREDMAILBOX)
    recipient.Resolve()
    # Find the folders needed
    inbox = outlookMAPI.GetSharedDefaultFolder(recipient, 6)
    parentfolder = inbox.Parent
    moveitfolder = parentfolder.Folders("Move IT Notifications")
    pendingfolder = moveitfolder.Folders("Pending")
    messages = inbox.Items
    for item in messages:
        # first check if the filename (email title) indicates a MoveIT request
        if item.subject.startswith("Continued Privileged Access required for"):
            process_mail(item, outlook, SHAREDMAILBOX, SIGNATURE)
            moveitsent = True
            # move the processed mail to the "pending" folder, awaiting reply
            item.UnRead = False
            item.Move(pendingfolder)
    # Feedback in case no MoveIT requests found in inbox
    if moveitsent == False:
        print("No MoveIT requests found.")
        
if __name__ == "__main__":
    main()
