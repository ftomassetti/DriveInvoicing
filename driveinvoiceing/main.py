from __future__ import print_function
import httplib2
import os

from apiclient import discovery
import oauth2client
from oauth2client import client
from oauth2client import tools
from apiclient.http import MediaIoBaseDownload
import io


try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/drive-python-quickstart.json
# https://www.googleapis.com/auth/documents
SCOPES = 'https://www.googleapis.com/auth/drive https://www.googleapis.com/auth/documents'
#CLIENT_SECRET_FILE = 'client_secret.json'
CLIENT_SECRET_FILE = 'secret1.json'
APPLICATION_NAME = 'Drive Invoicing'


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'drive-python-quickstart.json')

    store = oauth2client.file.Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials


def get_directory(service, folder_name):
    page_token = None
    res = []
    while True:
        response = service.files().list(q="mimeType = 'application/vnd.google-apps.folder' and name = '%s'" % folder_name,
                                     spaces='drive',
                                     fields='nextPageToken, files(id, name, properties)',
                                     pageToken=page_token).execute()
        for file in response.get('files', []):
            res.append(file)
        page_token = response.get('nextPageToken', None)
        if page_token is None:
            break;
    if len(res) != 1:
        raise Exception("Folder not found %s" % folder_name)
    return res[0]


def get_content(service, file_name, folder_id):
    page_token = None
    res = []
    while True:
        response = service.files().list(q=" '%s' in parents and name = '%s'" % (folder_id, file_name),
                                     spaces='drive',
                                     fields='nextPageToken, files(id, name, properties)',
                                     pageToken=page_token).execute()
        for file in response.get('files', []):
            res.append(file)
        page_token = response.get('nextPageToken', None)
        if page_token is None:
            break;
    if len(res) != 1:
        raise Exception("File not found %s (res=%s)" % (file_name, res))
    return res[0]


def get_folder(service, folder_name):
    page_token = None
    res = []
    while True:
        response = service.files().list(q="mimeType = 'application/vnd.google-apps.folder' and name = '%s'" % folder_name,
                                     spaces='drive',
                                     fields='nextPageToken, files(id, name, properties)',
                                     pageToken=page_token).execute()
        for file in response.get('files', []):
            res.append(file)
        page_token = response.get('nextPageToken', None)
        if page_token is None:
            break;
    if len(res) != 1:
        raise Exception("Folder not found %s (res=%s)" % (folder_name, res))
    return res[0]


def copy_file(service, original_id, new_file_name, folder_id):
    res = service.files().copy(fileId=original_id, body={"parents":[folder_id], "name":new_file_name}).execute()
    print(res)
    return res


def download_file_as_ooffice(service, file_id, file_name):
    download_file_as(service, file_id, 'application/vnd.oasis.opendocument.text', file_name)

def download_file_as_pdf(service, file_id, file_name):
    download_file_as(service, file_id, 'application/pdf', file_name)

def download_file_as(service, file_id, media_type, file_name):
    request = service.files().export_media(fileId=file_id, mimeType=media_type)
    #fh = io.BytesIO()
    fh = io.FileIO(file_name, mode='wb')
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
        print("Download %d%%." % int(status.progress() * 100))


def main():
    """Shows basic usage of the Google Drive API.

    Creates a Google Drive API service object and outputs the names and IDs
    for up to 10 files.
    """
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('drive', 'v3', http=http)

    folder_id = get_folder(service, 'DriveInvoicing')['id']
    template_id = get_content(service, 'Template', folder_id)['id']
    print("template id %s" % template_id)

    invoice = {'number':20, 'date':{'day':11, 'month':'April', 'year':2016}, 'noVAT':True,
               'client':{'name':'Pinco', 'address':'Via Foo','vatID':'FOO123','contact':'Mr. Pallo'},
               'lines':[{'description':'Stuff done', 'amount':128.34, 'vatRate':20.0},
               {'description':'Other Stuff', 'amount':80.0, 'vatRate':20.0},
               {'description':'Third line', 'amount':85.0, 'vatRate':20.0}]}

    invoice_doc_id = copy_file(service, template_id, 'Invoice_%i' % invoice['number'], folder_id)['id']

    script_service = discovery.build('script', 'v1', http=http)
    request = {"function": "insertData", "devMode": True, "parameters": [
        invoice_doc_id, invoice['number'], invoice['date'], invoice['noVAT'], invoice['client'], invoice['lines']]}
    SCRIPT_ID = 'MQVDLm6PA-4WSCs1YNzsXtfeXYCx8nRBo'
    response = script_service.scripts().run(body=request, scriptId=SCRIPT_ID).execute()
    #print(response)
# 11,
#              {'day':11, 'month':'April', 'year':2016},
#              false,
#              {'name':'Pinco', 'address':'Via dei pinchi palli','vatID':'FOO1234','contact':'Mr. Pallo'},
#              [{'description':'Stuff done', 'amount':128.34, 'vatRate':20.0},
#               {'description':'Other Stuff', 'amount':80.0, 'vatRate':20.0},
#               {'description':'Third line', 'amount':85.0, 'vatRate':20.0}]);

    download_file_as_pdf(service, invoice_doc_id, 'Invoice_%i.pdf' % invoice['number'])
    #content = service.files().get_media(fileId=template_id).execute()
    #print(content)

if __name__ == '__main__':
    main()