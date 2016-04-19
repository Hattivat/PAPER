# -*- coding: utf-8 -*-
"""
Created on Wed Mar  9 10:51:59 2016

@author: Jan Janiczek
"""

import win32com.client as client

SHAREDMAILBOX = "xxxxxxxxxx"
TRASHFOLDER = "Untitled"

def clean_mail(mailbox=SHAREDMAILBOX, trashname=TRASHFOLDER):
    """Takes a mailbox and moves any PAMSNoReply mails found in its inbox to
    the trash folder of the provided name.
    
    takes:
    mailbox: a string containing the name of shared mailbox to be cleaned
    trashname: the name of the folder used for NoReply mail on that mailbox
    
    returns: nothing
    prints: the amount of emails moved (default 0)"""
    
    outlook = client.Dispatch("Outlook.Application")
    outlookMAPI = outlook.GetNamespace("MAPI")
    recipient = outlookMAPI.CreateRecipient(mailbox)
    recipient.Resolve()
    #Find the folders needed
    inbox = outlookMAPI.GetSharedDefaultFolder(recipient, 6)
    parentfolder = inbox.Parent
    trashfolder = parentfolder.Folders(trashname)
    messages = inbox.Items
    #Only used comprehension to see if it would speed things up; no noticeable
    #difference, kept it out of inertia.
    noreplymsgs = [msg for msg in messages if msg.SenderName == "PAMSNoReply@gsk.com"]
    counter = 0
    for item in noreplymsgs:
        item.UnRead = False
        item.Move(trashfolder)
        counter += 1
    print("{0} noreply emails moved.".format(counter))
    
if __name__ == "__main__":
    clean_mail()
    