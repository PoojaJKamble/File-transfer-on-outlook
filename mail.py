import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

mail_content = '''Greetings,
Please find attached Report.


(This mail is sent using python script)

Thank You!'''


#The mail addresses and password
sender_address = "pooja_kamble@hotmail.com"
sender_pass = "*********"
receiver_address = "pooja_kamble@hotmail.com"
#Setup the MIME
message = MIMEMultipart()
message['From'] = sender_address
message['To'] = receiver_address
message['Subject'] = 'Report'   #The subject line

#The body and the attachments for the mail
message.attach(MIMEText(mail_content, 'plain'))


# open the file to be sent
filename = "07-25-2021Billing.xlsx"
attachment = open(r"C:\Users\pkamble7\OneDrive - DXC Production\PycharmProjects\BillingFolder\07-25-2021Billing.xlsx", "rb")

# instance of MIMEBase and named as p
p = MIMEBase('application', 'octet-stream')

# To change the payload into encoded form
p.set_payload((attachment).read())

# encode into base64
encoders.encode_base64(p)

p.add_header('Content-Disposition', "attachment; filename= %s" % filename)

# attach the instance 'p' to instance 'msg'
message.attach(p)


#Create SMTP session for sending the mail
session = smtplib.SMTP(host='smtp.office365.com', port=587) #Outgoining Mail Server
session.starttls() #enable security
session.login(sender_address, sender_pass) #login with mail_id and password
text = message.as_string()
session.sendmail(sender_address, receiver_address, text)

session.quit()
print('Report Sent Successfully!')

