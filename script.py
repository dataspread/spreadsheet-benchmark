'''
Original template from https://medium.com/@victor.perez.berruezo/execute-google-apps-script-functions-or-sheets-macros-programmatically-using-python-apps-script-ec8343e29fcd
For more help https://developers.google.com/apps-script/api/quickstart/python
'''

import pickle
import os.path
from googleapiclient import errors
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

import time
import random

# If scopes are not sufficient, go to File > Project Properties > Scopes of 
# the Apps Scripts project and copy all Scopes
# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    
def get_scripts_service():
    """Calls the Apps Script API.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            # Credentials path from the credentials .json file 
            # TODO: Replace if file path or file name changes
            flow = InstalledAppFlow.from_client_secrets_file(
                './client_id.json', SCOPES) 
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    return build('script', 'v1', credentials=creds)


service = get_scripts_service()
# API ID from Apps Scripts API Executable
# TODO: Replace for every new API/project
API_ID = "API_ID" 

# TODO: Set test to false once confirmed that the API works for easier debugging
# Replace sizes with the dataset sizes corresponding to your project
test = True
if test:
    unique_sizes = [10000, 20000, 30000, 40000, 50000, 60000, 70000, 80000, 90000]
    sizes = []
    for size in unique_sizes:
        sizes += [size] * 10
else:
    sizes = [10000]

random.shuffle(sizes)
print("Number of Trials: ", len(sizes))
print("sizes: ", sizes)

for i, size in enumerate(sizes):
    # TODO: Replace `function_name` with function to be called in Apps Script
    # Parameters as arguments to the function 
    request = {
        "function": "function_name",
        "parameters": [size]
        } 
        
    try:
        response = service.scripts().run(body=request, scriptId=API_ID).execute()
        print("Trial {}; Size {}: {}".format(i, size, response))
        # wait at least 10 seconds for next request to not overload api
        time.sleep(10)
    except errors.HttpError as error:
        # The API encountered a problem.
        print(error.content)