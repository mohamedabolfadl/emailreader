"""

Script to read emails from gmail and download the attachments

input:
    - email and password
    - target_senders the email addresses expected to send
    - from_date to_date: The limit of dates
output:
    02_data/output/uber_receipts.csv    

"""


#-- Clear variables
from IPython import get_ipython
ipython = get_ipython()
ipython.magic("%reset  -sf")




import imaplib
import base64
import os
import email
import re


#-- Make sure we are at root of project
os. chdir("C:\\Users\\Mohamed Ibrahim\\Box Sync\\bot\\emailreader")

target_senders = ['kundenservice@free-now.com','help@mytaxi.com']

from_date = "01-Apr-2019"
to_date = "07-Jul-2019"



#-- Set date into format of year month day
def getFormattedDate(email_message):
    try:
        dt = email_message['date'][5:16].replace(" ","_")
        dt = dt.replace('Jan','01').replace('Feb','02').replace('Mar','03').replace('Apr','04').replace('May','05').replace('Jun','06')
        dt = dt.replace('Jul','07').replace('Aug','08').replace('Sep','09').replace('Oct','10').replace('Nov','11').replace('Dec','12')
        dt = dt[6:]+"_"+dt[3:5]+"_"+dt[0:2]
    except:
        dt = ""
        pass
    return dt

#-- Extract sender email from text
def getFormattedSender(email_message):
    try:
        em = email_message["from"]
        sender = em[em.find("<")+1:em.find(">")]
    except:
        sender = ""
        pass
    return sender


#-- Get email credentials
email_user = input(['Enter Email'])
email_pass = input(['Enter password'])


#-- Create connection instance
mail = imaplib.IMAP4_SSL("imap.gmail.com",993)

#-- Log in
mail.login(email_user, email_pass)

#-- Focus on Inbox
mail.select('Inbox')



#-- Loop over the senders
senders_i = 0

while senders_i< (len(target_senders)):

    #-- Filter date and sender
    tp, data = mail.search(None, '(SINCE "'+from_date+'" BEFORE "'+ to_date+'")', '(FROM "'+target_senders[senders_i]+'")')
    
    
    
    mail_ids = data[0]
    id_list = mail_ids.split()
    for num in data[0].split():
        typ, data = mail.fetch(num, '(RFC822)' )
        raw_email = data[0][1]
    # converts byte literal to string removing b''
        try:
            raw_email_string = raw_email.decode('utf-8')
            email_message = email.message_from_string(raw_email_string)
            last_email_message = email_message
        except:
            email_message = last_email_message
            pass
    
        #-- Get email date
        dt = getFormattedDate(email_message)
        #-- Get sender
        sender = getFormattedSender(email_message)
    
    
  
        
    # downloading attachments
        for part in email_message.walk():
            # this part comes from the snipped I don't understand yet... 
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue
            fileName = part.get_filename()
            
            #-- Add date and sender
            fileName = dt+"__"+sender+"__"+fileName
            
            
            if bool(fileName):
                filePath = os.path.join('02_data/output/', fileName)
                fp = open(filePath, 'wb')
                fp.write(part.get_payload(decode=True))
                fp.close()
    senders_i = senders_i + 1
             