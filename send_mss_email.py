#!/usr/local/bin/python3
import smtplib
from email.mime.text import MIMEText
import ast
from email_auth import sender_email, sender_password

def getUsers():
    with open('./values.txt', 'r') as f:
        contents = f.read()
        lines = contents.splitlines()
        f.close()
        #print(lines)
        return lines

def getMessage(message_type):
    with open('/Users/cyee/Documents/mss_registrations/{}_message.txt'.format(message_type), 'r') as f:
        contents = f.read()
        contents = str(contents)
        #print(contents)
        f.close()
        return contents

def email():
    lines = getUsers()
    for line in lines:
        values = ast.literal_eval(line)
        first = values['first']
        last = values['last']
        city = values['city']
        email = values['email']
        use = values['use']
        password = values['password']
        username = values['username']
        lower_last = last.lower()
        studentusername = 'student{}'.format(lower_last)

        msg_content = ''
        if use == 'Use':
            contents = getMessage('user') 
            msg_content = contents.format(first=first, password=password, username=username, studentusername=studentusername)
        else:
            contents = getMessage('preview') 
            msg_content = contents.format(first=first, password=password, username=username)

        message = MIMEText(msg_content, 'html')

        message['From'] = 'Cameron Yee  <{}>'.format(sender_email)
        message['To'] = '{name} <{email}>'.format(name=first + last, email=email)
        message['Subject'] = 'BSCS Middle School Science'

        msg_full = message.as_string()

        server = smtplib.SMTP('smtp.office365.com:587')
        server.starttls()
        server.login('{}'.format(sender_email), '{}'.format(sender_password))
        server.sendmail('{}'.format(sender_email),
                        [email, sender_email], #send to self as confirmation
                        msg_full)
        server.quit()


if __name__ == '__main__':
    email()
