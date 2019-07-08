"""

Script to read emails from outlook and store emails in a csv file

input:
    NA
    
output:
    02_data/output/uber_receipts.csv    

"""

#-- Clear variables
from IPython import get_ipython
ipython = get_ipython()
ipython.magic("%reset  -sf")

#-- Import libraries
import win32com.client
import os
import re
import numpy as np
import pandas as pd


#-- Display settings
pd.set_option('display.max_rows', 500)
pd.set_option('display.max_columns', 10)
pd.set_option('display.width', 1000)

#-- keyword
keyword = "Trip Fare"


#-- Initialize dataframe
df = pd.DataFrame(columns=['date', 'sender', 'body','amount','fromAddress','destinationAddress','fromTime','toTime'])


#-- Connect to outlook
outlook=win32com.client.Dispatch("Outlook.Application").GetNameSpace("MAPI")

#-- Get inbox folder
# Inbox 6
# Sent 5
inbox=outlook.GetDefaultFolder(6) #Inbox default index value is 6

#-- Get all inbox items
messages = inbox.Items

#-- Email counter
i=0
for msg in messages:

    #-- Get minimal message components
    try:
        sender=msg.Sender
        date=msg.senton.date()
        body=msg.body        
    except:
            #print("Error")
            continue
    #-- Make sure it contains the info you need
    try:
        if bool(re.search(keyword ,body)):
            #-- Get trip amount
            amount = re.findall('\d+\.\d+',re.findall('Total \\t\$\d+\.\d+', body)[0])[0]
            
            #-- Get start and destination
            tripTrack = re.findall('\d+\:\d+.+Egypt',body.replace('\n',''))[0]
            firstIndexOfEgypt = tripTrack.find('Egypt') 

            fromAddress = tripTrack[1:5+firstIndexOfEgypt ].replace('\t','').replace('\r','')
            destinationAddress = re.findall('\d+\:\d+.+Egypt',tripTrack[firstIndexOfEgypt+5:])[0].replace('\t','').replace('\r','')
            
            fromTime = re.findall('\d+:\d+',fromAddress)[0]
            toTime = re.findall('\d+:\d+',destinationAddress)[0]


            #-- Insert into dataframe
            df.loc[i]=list([date,sender,body,amount,fromAddress,destinationAddress,fromTime,toTime])

            i=i+1
            #print("Message on "+date+" from "+sender)
            #print("Added "+i+" emails")
            
    except:
       continue



#-- Dump as csv
df.loc[:,['date','amount','fromAddress','destinationAddress', 'fromTime', 'toTime']].to_csv('02_data/output/uber_receipts.csv')



"""
if False:
    
    import pdfkit
    path_wkthmltopdf = r'C:\python37\wkhtmltopdf\bin\wkhtmltopdf.exe'
    config = pdfkit.configuration(wkhtmltopdf=path_wkthmltopdf)
    pdfkit.from_string(df.loc[1,'body'], 'out.pdf', configuration=config)
    
    
    
    
    html_str = df.loc[1,'body']
    Html_file= open("test.html","w")
    Html_file.write(html_str)
    Html_file.close()
    
    
    
    pdfkit.from_url("http://google.com", "out.pdf", configuration=config)
    
    import pdfkit 
    import wkhtmltopdf
    pdfkit.from_string('Hello!', 'out.pdf')
    
    
    import pdfkit
    path_wkthmltopdf = r'C:\Python27\wkhtmltopdf\bin\wkhtmltopdf.exe'
    config = pdfkit.configuration(wkhtmltopdf=path_wkthmltopdf)
    pdfkit.from_url("http://google.com", "out.pdf", configuration=config)
    
    
    
    pdfkit.from_url('http://google.com', 'out.pdf')
    
    
    from weasyprint import HTML, CSS
    
    HTML(df.loc[:,'body']).write_pdf("/path_for_storing_your_PDF/file_name.pdf",  
    stylesheets=[CSS(string="body { font-color: red }")])
    
    
    
    
    body = df.iloc[1,2].replace('\n','')
    reg = '\d+\:\d+.+Egypt'
    
    tripTrack = re.findall('\d+\:\d+.+Egypt',body)[0]
    firstIndexOfEgypt = tripTrack.find('Egypt') 
    
    fromAddress = tripTrack[1:5+firstIndexOfEgypt ].replace('\t','').replace('\r','')
    destinationAddress = re.findall('\d+\:\d+.+Egypt',tripTrack[firstIndexOfEgypt+5:])[0].replace('\t','').replace('\r','')
    
    
    
    secondIndexOfEgypt = tripTrack[firstIndexOfEgypt+5:].find('Egypt') 
    destinationAddress = tripTrack[1:5+firstIndexOfEgypt ].replace('\t','').replace('\r','')
    
    
    #\t08:58am \t\r\n40 Ahmed Fakhry, Al Manteqah as Sadesah, Nasr City, Cairo Governorate, Egypt
    
    
    
    
                    
    
    bool(bd.find("stss"))
        
    print(msg.Sender)
    
    bool(re.search("stss",bd))
        
        
        
    message2=message.GetLast()
    subject=message2.Subject("XID Creation Report")
    body=message2.body
    date=message2.senton.date()
    sender=message2.Sender
    attachments=message2.Attachments
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    import win32com.client
    import win32com
    import os
    import sys
    
    f = open("02_data/output/testfile.txt","w+")
    
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    accounts= win32com.client.Dispatch("Outlook.Application").Session.Accounts;
    
    
    def emailleri_al(folder):
        messages = folder.Items
        a=len(messages)
        if a>0:
            for message2 in messages:
                 try:
                    sender = message2.SenderEmailAddress
                    if sender != "":
                        print(sender, file=f)
                 except:
                    print("Error")
                    print(account.DeliveryStore.DisplayName)
                    pass
    
                 try:
                    message2.Save
                    message2.Close(0)
                 except:
                     pass
    
    
    
    for account in accounts:
        global inbox
        inbox = outlook.Folders(account.DeliveryStore.DisplayName)
        print("****Account Name**********************************",file=f)
        print(account.DisplayName,file=f)
        print(account.DisplayName)
        print("***************************************************",file=f)
        folders = inbox.Folders
    
        for folder in folders:
            print("****Folder Name**********************************", file=f)
            print(folder, file=f)
            print("*************************************************", file=f)
            emailleri_al(folder)
            a = len(folder.folders)
    
            if a>0 :
                global z
                z = outlook.Folders(account.DeliveryStore.DisplayName).Folders(folder.name)
                x = z.Folders
                for y in x:
                    emailleri_al(y)
                    print("****Folder Name**********************************", file=f)
                    print("..."+y.name,file=f)
                    print("*************************************************", file=f)
    
    
    
    print("Finished Succesfully")
    
    
    
    import win32com.client
    from win32com.client import Dispatch, constants
    
    const=win32com.client.constants
    olMailItem = 0x0
    obj = win32com.client.Dispatch("Outlook.Application")
    newMail = obj.CreateItem(olMailItem)
    newMail.Subject = "I AM SUBJECT!!"
    # newMail.Body = "I AM\nTHE BODY MESSAGE!"
    newMail.BodyFormat = 2 # olFormatHTML https://msdn.microsoft.com/en-us/library/office/aa219371(v=office.11).aspx
    newMail.HTMLBody = "<HTML><BODY>Enter the <span style='color:red'>message</span> text here.</BODY></HTML>"
    newMail.To = "email@demo.com"
    attachment1 = r"C:\Temp\example.pdf"
    newMail.Attachments.Add(Source=attachment1)
    newMail.display()
    newMail.send()
    
    
    
    
    import win32com.client
    import os
    outlook=win32com.client.Dispatch("Outlook.Application").GetNameSpace("MAPI")
    inbox=outlook.GetDefaultFolder(6) #Inbox default index value is 6
    # Inbox 6
    # Sent 5
    message=inbox.Items
    message2=message.GetLast()
    subject=message2.Subject("XID Creation Report")
    body=message2.body
    date=message2.senton.date()
    sender=message2.Sender
    attachments=message2.Attachments
    print(subject)
    print(body)
    print(sender)
    print(attachments.count)
    print(date)
    
    
    #pip install pywin32
    #pip install pypiwin32
    #pip install win32api
    #pip install win32com
    
    
    import pywin32
    import win32com.client
"""