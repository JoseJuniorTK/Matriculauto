import email
from email import policy
from email.parser import BytesParser
import html2text
import re

filepath = './Demanda.eml'

with open(filepath) as email_file:
    email_message = email.message_from_file(email_file)

if email_message.is_multipart():
    for part in email_message.walk():
        message = str(part.get_payload(decode=True))
        plain_message = html2text.html2text(message)
        # plain_message = message
        # print(plain_message)
        print()

plain_message = plain_message.encode('utf8', 'strict')
plain_message2 = plain_message.decode('utf8', 'strict')

plain_message3 = re.sub('[^A-Za-z0-9]+', '', plain_message2)
# plain_message = html2text.html2text(message)

with open("Output.txt", "w", encoding='ascii') as text_file:
    text_file.write(plain_message3)
