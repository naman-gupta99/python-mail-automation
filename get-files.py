import base64
import email
import imaplib
import os
import config


def main():
    email_user = config.USERNAME
    email_password = config.PASSWORD

    mail = imaplib.IMAP4_SSL('imap.gmail.com', 993)
    mail.login(email_user, email_password)

    mail.select('"Axis Bank"')

    typ, data = mail.search(None, 'ALL')
    mail_ids = data[0]
    id_list = mail_ids.split()

    for num in id_list:
        typ, data = mail.fetch(num, '(RFC822)')

        raw_email = data[0][1]

        raw_email_string = raw_email.decode('utf-8')
        email_message = email.message_from_string(raw_email_string)

        for part in email_message.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue
            fileName = part.get_filename()
            if bool(fileName):
                filePath = os.path.join(
                    'C:/Users/kungf/Desktop/Statement/Files/Downloads', fileName)
                if not os.path.isfile(filePath):
                    fp = open(filePath, 'wb')
                    fp.write(part.get_payload(decode=True))
                    fp.close()


if __name__ == '__main__':
    main()
