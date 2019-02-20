import os
import pickle
from googleapiclient import discovery

if os.path.exists('token.pickle'):
    cred = 'token.pickle'
else:
    cred = 'sheets/token.pickle'

with open(cred, 'rb') as token:
    credentials = pickle.load(token)

service = discovery.build('sheets', 'v4', credentials=credentials)
