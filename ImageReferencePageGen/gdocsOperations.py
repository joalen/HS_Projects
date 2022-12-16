import time 
from progress.bar import IncrementalBar
import os 
from gdoctableapppy import gdoctableapp
import secrets, string


def createTable(oauthCreds, docID, **kwargs): 
    """
    Creates a table in a Google Doc. The table has two columns and the number of rows is equal to the sum of the values in the list "imageRepeats". 
    The first column is a unique identifier and the second column is a string that contains the chapter number and page number.

    :param ```oauthCreds```  - Google JSON API File unique to your account to process the API requests 
    :param ```docID```  - Google Docs ID. Found from the URL of a Google Doc
    :param other: Other Parameters required for this function: 
        •  imageRepeats - an ```int``` array of repeating images on a single page (each index is a page) \n
        •  imagesDir - a ```str``` that represents the folder directory where the images are found \n
        •  end_ - an ```int``` that represents the page limit (including). \n
        •  strt - an ```int``` that represents the starting page (including). \n 
        •  chapterNumber - an ```int``` that represents the chapter number of a given PDF or text. \n
    
    :return - a ```list``` that contains replaceable strings to add the images. These are also ordered! Do not modify!
    """

    resource = { 
    "oauth2": oauthCreds,
    "documentId": docID,
    "rows": sum(kwargs["imageRepeats"]),
    "columns": 2,
    "createIndex": 1,
    }

    uniqueIdentifiers = []
    valuesToTable = []
    counter = 0
    current_imageRepeatsIndex = 0 

    for i in range(len(os.listdir(kwargs["imagesDirectory"]))): 
        uniqueIdentifiers.append(''.join(secrets.choice(string.ascii_uppercase + string.ascii_lowercase) for i in range(25)))
    

    bar_inst_0 = IncrementalBar('Setting Table ', max=(kwargs["end_"] + 1 - kwargs["strt"] + 1), suffix='%(percent).1f%% - %(eta)ds')  
    for i in range(kwargs["strt"], kwargs["end_"] + 1):
        for j in range(kwargs["imageRepeats"][current_imageRepeatsIndex]):
            chapterNumber = kwargs["chapterNumber"]
            valuesToTable.append([f"--sample-- {uniqueIdentifiers[counter]}", f"Chapter {chapterNumber} - pg. {i}"])
            counter += 1
            bar_inst_0.next()
            time.sleep(1) 
        current_imageRepeatsIndex += 1

    resource["values"] = valuesToTable
    gdoctableapp.CreateTable(resource)

    return uniqueIdentifiers

def appendTable(oauthCreds, docID, **kwargs):
    """
    This function appends a row to a table in a Google Doc. If a table is not found, the function will invoke ```createTable()``` to handle original task.
    
    :param ```oauthCreds```: the credentials for the Google Drive API
    :param ```docID```: the document ID of the Google Doc you want to append to

    :param other: Other Parameters required for this function: 
        •  imageRepeats - an ```int``` array of repeating images on a single page (each index is a page) \n
        •  imagesDir - a ```str``` that represents the folder directory where the images are found \n
        •  end_ - an ```int``` that represents the page limit (including). \n
        •  strt - an ```int``` that represents the starting page (including). \n 
        •  chapterNumber - an ```int``` that represents the chapter number of a given PDF or text. \n
    """

    DOCUMENTID = docID

    resource = { 
    "oauth2": oauthCreds,
    "documentId": DOCUMENTID,
    "tableIndex": 0,
    }

    uniqueIdentifiers = []
    valuesToTable = []
    counter = 0
    current_imageRepeatsIndex = 0 

    for i in range(len(os.listdir(kwargs["imagesDirectory"]))): 
        uniqueIdentifiers.append(''.join(secrets.choice(string.ascii_uppercase + string.ascii_lowercase) for i in range(25)))
    

    bar_inst_0 = IncrementalBar('Setting Table ', max=(kwargs["end_"] + 1 - kwargs["strt"] + 1), suffix='%(percent).1f%% - %(eta)ds')  
    for i in range(kwargs["strt"], kwargs["end_"] + 1):
        for j in range(kwargs["imageRepeats"][current_imageRepeatsIndex]):
            chapterNumber = kwargs["chapterNumber"]
            valuesToTable.append([f"--sample-- {uniqueIdentifiers[counter]}", f"Chapter {chapterNumber} - pg. {i}"])
            counter += 1
            bar_inst_0.next()
            time.sleep(1) 
        current_imageRepeatsIndex += 1
    try:
        gdoctableapp.AppendRow(resource)
    except ValueError: 
        uniqueIdentifiers = createTable(oauthCreds, docID, imageRepeats=kwargs["imageRepeats"], strt=kwargs["strt"], end_=kwargs["end_"], chapterNumber=kwargs["chapterNumber"], imagesDirectory=kwargs["imagesDirectory"])

def checkGoogleOAuth(googleScopes): 
    """
    If there's a token.json file, use it. If not, use the client_secrets.json file to create a
    token.json file
    
    :param googleScopes: The scopes you want to use for your Google API
    :return: The generated credentials for the user.

    """
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request

    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', googleScopes)
    
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'client_secrets.json', googleScopes)
            creds = flow.run_local_server(port=0)
        # Save the credentials to avoid redundancy 
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    
    return creds
