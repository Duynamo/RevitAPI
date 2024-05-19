import os
from email.message import EmailMessage
import ssl
import smtplib

emailSender = 'duyvultv@gmail.com'
emailPassword = 'nmgx vjtz eshc tbvu'
#emailReceiver = 'vudinhduybm@gmail.com'
emailReceivers = ['vudinhduybm@gmail.com', 'duyvultv@gmai.com', 'vuducvubm123@gmail.com']
emailCCList = []
subject = 'Remind sending PJ Status to FUSO customers'
bodyList = ["""
You need to send an email to FUSO customers to inform the status of the following PJ:
PJ_1, PJ_2.
This email is automatically sent, no reply.
""",
"""
You need to send an email to FUSO customers to inform the status of the following PJ:
PJ_14, PJ_24.
This email is automatically sent, no reply.
""",
"""
You need to send an email to FUSO customers to inform the status of the following PJ:
PJ_12, PJ_2222.
This email is automatically sent, no reply.
"""]

#body = "\n".join(bodyList)
body = [body for body in bodyList]


#region ___send to multiple Receiver with the same email
# em = EmailMessage()
# em['From'] = emailSender
# em['To'] = [emailReceiver for emailReceiver in emailReceivers]
# em['Subject'] = subject
# em.set_content(body)

# context = ssl.create_default_context()

# with smtplib.SMTP_SSL('smtp.gmail.com', 465, context = context) as smtp:
#     smtp.login(emailSender, emailPassword)
#     smtp.sendmail(emailSender, emailReceivers, em.as_string())

#endregion

context = ssl.create_default_context()
with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
    smtp.login(emailSender, emailPassword)
    for emailReceiver, body in zip(emailReceivers, bodyList):
        em = EmailMessage()
        em['From'] = emailSender
        em['To'] = emailReceiver
        em['Subject'] = subject
        em.set_content(body)
        
        smtp.sendmail(emailSender, emailReceiver, em.as_string())