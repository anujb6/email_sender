import pandas as pd
import smtplib

my_name = "anuj bhor"
my_email = "ravabhoir3@gmail.com"
my_password = "anujb109$"

server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server.ehlo()
server.login(my_email, my_password)

email_list = pd.read_excel("excel sheet name");

#getting all the info inside excel sheet
all_names = email_list['Name']
all_emails = email_list['Email']
all_subject = email_list['Subject']
all_messages = email_list['Message']

#looping
for idx in range(len(all_emails)):

    name = all_names[idx]
    email = all_emails[idx]
    subject = all_subject[idx]
    message = all_messages[idx]

    #creating a email to send
    email_struct = ("From: {0} <{1}>\n"
                  "To: {2} <{3}>\n"
                  "Subject: {4}\n\n"
                  "{5}"
                  .format(my_name, my_email, name, email, subject, message))


    try:
        server.sendmail(my_email, [email], email_struct)
        print('Email to {} succesfully sent!\n\n'.format(email))
    except Exception as e:
        print('Email to {} could not be sent, Because {}\n\n'.format(email, str(e)))

server.close()                          