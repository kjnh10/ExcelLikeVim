import authorization

from httplib2 import Http
from apiclient.discovery import build
import numpy as np
from xlwings import Workbook, Range

SCOPES = 'https://www.googleapis.com/auth/gmail.modify'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Gmail'

def test():
    credentials = authorization.get_credentials(SCOPES,CLIENT_SECRET_FILE,APPLICATION_NAME)
    service = build('gmail', 'v1', http=credentials.authorize(Http()))
    results = service.users().labels().list(userId='me').execute()
    labels = results.get('labels', [])

    if not labels:
        print 'No labels found.'
    else:
        print 'Labels:'
        i = 1
        for label in labels:
            try:
                wb = Workbook('Book1')
                Range((i,1)).value = label['name'].encode("shift-jis")
            except:
                print "error"
            i += 1

def rand_numbers():
    """ produces standard normally distributed random numbers with shape (n,n)"""
    wb = Workbook('Book1')  # Creates a reference to the calling Excel file
    n = int(Range('Sheet1', 'B1').value)  # Write desired dimensions into Cell B1
    rand_num = np.random.randn(n, n)
    Range('Sheet1', 'C3').value = rand_num

