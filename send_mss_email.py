import smtplib
from email.mime.text import MIMEText
import ast

def email_auth():
    with open('./email_auth.txt', 'r') as f:
        contents = f.read()
        lines = contents.splitlines()
        sender_email = lines[0]
        sender_password  = lines[1]
        return sender_email, sender_password

def getUsers():
    with open('./values.txt', 'r') as f:
        contents = f.read()
        lines = contents.splitlines()
        f.close()
        #print(lines)
        return lines

def getMessage():
    with open('/Users/cyee/Desktop/mss_registrations/message.txt', 'r') as f:
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

        contents = getMessage() 
        msg_content = contents.format(first=first, last=last, password=password, username=username, studentusername=studentusername)
        message = MIMEText(msg_content, 'html')

        sender_email = email_auth()[0]
        sender_password = email_auth()[1]

        message['From'] = 'Cameron Yee  <{}>'.format(sender_email)
        message['To'] = '{name} <{email}>'.format(name=first + last, email=email)
        message['Subject'] = 'BSCS Middle School Science'

        msg_full = message.as_string()

        server = smtplib.SMTP('smtp.office365.com:587')
        server.starttls()
        server.login('{}'.format(sender_email), '{}'.format(sender_password))
        server.sendmail('{}'.format(sender_email),
                        [email],
                        msg_full)
        server.quit()


if __name__ == '__main__':
    email()
