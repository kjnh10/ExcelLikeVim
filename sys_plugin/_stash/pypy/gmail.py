# -*- coding: utf-8 -*-
"""Send an email message from the user's account.
"""
import authorization
from apiclient.discovery import build
from httplib2 import Http

import base64
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import mimetypes
import os

from apiclient import errors

SCOPES = 'https://www.googleapis.com/auth/gmail.modify'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Gmail'

def SendMessage(service, user_id, message):# {{{
  """Send an email message.

  Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    message: Message to be sent.

  Returns:
    Sent Message.
  """
  try:
    message = (service.users().messages().send(userId=user_id, body=message)
               .execute())
    print 'Message Id: %s' % message['id']
    return message
  except errors.HttpError, error:
    print 'An error occurred: %s' % error# }}}

def CreateMessage(sender, to, cc, bcc, subject, message_text):# {{{
  """Create a message for an email.

  Args:
    sender: Email address of the sender.
    to: Email address of the receiver.(Case multi address:comma separated.)
    cc: Email address of the receiver.(Case multi address:comma separated.)
    bcc: Email address of the receiver.(Case multi address:comma separated.)
    subject: The subject of the email message.
    message_text: The text of the email message.

  Returns:
    An object containing a base64 encoded email object.
  """
  message = MIMEText(message_text)
  message['to'] = to
  # message['cc'] = cc
  # message['bcc'] = bcc
  message['from'] = sender
  message['subject'] = "=?utf-8?B?" + base64.b64encode(subject) +"?="
  #message['subject'] = subject
  return {'raw': base64.b64encode(message.as_string())}# }}}

def CreateMessageWithAttachment(sender, to, subject, message_text, file_dir, filename):# {{{ 
  """Create a message for an email.
  Args:
    sender: Email address of the sender.
    to: Email address of the receiver.
    subject: The subject of the email message.
    message_text: The text of the email message.
    file_dir: The directory containing the file to be attached.
    filename: The name of the file to be attached.

  Returns:
    An object containing a base64 encoded email object.
  """
  message = MIMEMultipart()
  message['to'] = to
  message['from'] = sender
  message['subject'] = subject

  msg = MIMEText(message_text)
  message.attach(msg)

  path = os.path.join(file_dir, filename)
  content_type, encoding = mimetypes.guess_type(path)

  if content_type is None or encoding is not None:
    content_type = 'application/octet-stream'

  main_type, sub_type = content_type.split('/', 1)
  if main_type == 'text':
    fp = open(path, 'rb')
    msg = MIMEText(fp.read(), _subtype=sub_type)
    fp.close()
  elif main_type == 'image':
    fp = open(path, 'rb')
    msg = MIMEImage(fp.read(), _subtype=sub_type)
    fp.close()
  elif main_type == 'audio':
    fp = open(path, 'rb')
    msg = MIMEAudio(fp.read(), _subtype=sub_type)
    fp.close()
  else:
    fp = open(path, 'rb')
    msg = MIMEBase(main_type, sub_type)
    msg.set_payload(fp.read())
    fp.close()

  #msg.add_header('Content-Disposition', 'attachment', filename=filename)
  msg.add_header('Content-Disposition', 'inline', filename=filename) #inlineにすると埋め込まれる｡
  message.attach(msg)
  return {'raw': base64.b64encode(message.as_string())}# }}}

def SendSimpleMail(to, subject, body, user_id="me", cc="", bcc=""):# {{{
    subject = subject.encode("utf-8")
    body = body.encode("utf-8")
    credentials = authorization.get_credentials(SCOPES,CLIENT_SECRET_FILE,APPLICATION_NAME)
    service = build('gmail', 'v1', http=credentials.authorize(Http()))
    SendMessage(service, user_id, CreateMessage(user_id, to, cc, bcc, subject, body))# }}}

if __name__ == '__main__':
    subject = u' タイトル '
    body = u"""
    幸治
    渡邉
    """
    SendSimpleMail('koji0708phone@gmail.com, koji07082002@gmail.com', subject, body)

# if __name__ == '__main__':
#     credentials = authorization.get_credentials(SCOPES,CLIENT_SECRET_FILE,APPLICATION_NAME)
#     service = build('gmail', 'v1', http=credentials.authorize(Http()))
#
#     mail = CreateMessageWithAttachment('me','koji0708phone@gmail.com','title','honbunn','C:\\Users\\bc0074854\\Desktop','2014-03-15 14.52.33.jpg')
#
#     SendMessage(service, 'me', mail)# }}}

