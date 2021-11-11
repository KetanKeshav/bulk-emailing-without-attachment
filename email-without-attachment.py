# bulk emailing without attachment

import smtplib
import xlrd
from email.message import EmailMessage

server=smtplib.SMTP_SSL("smtp.gmail.com",465)
server.login("insert_email@gmail.com","insert_pwd")


#logic to get the names and emails list from excel sheet
names=list()
emails=list()
path=input("Enter Path of Excel File\n")
wb = xlrd.open_workbook(path) 
sheet = wb.sheet_by_index(0)
R=sheet.nrows

for i in range(R): 
    #Column 1 - Names
    name=sheet.cell_value(i,0)
    
    #Column 2 - Corresponding email address
    email=sheet.cell_value(i,1)
    
    names.append(name)
    emails.append(email)

print("names and emails extracted sucessfully")

for i in range(len(names)):
    msg=EmailMessage()

    msg['Subject']="insert-subject"
    msg['From']='insert-email'
    msg['To']=emails[i]
    body="""
Dear {0},
Greetings!

Thanks & Regards,
Name
""".format(names[i])

    msg.set_content(body)
    server.send_message(msg)
    
server.quit()
print("Smile its sent successfully!!")