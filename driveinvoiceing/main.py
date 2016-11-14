from __future__ import print_function
import httplib2
import os

from apiclient import discovery
import oauth2client
from oauth2client import client
from oauth2client import tools
from apiclient.http import MediaIoBaseDownload
import io
import simplejson
import argparse

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/drive-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/drive https://www.googleapis.com/auth/documents'
CLIENT_SECRET_FILE = 'secret.json'
APPLICATION_NAME = 'Drive Invoicing'
SCRIPT_ID = 'MQVDLm6PA-4WSCs1YNzsXtfeXYCx8nRBo'

DEFAULT_CURRENCY = 'Euro'

flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args(args=[])


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
                                   'drive_invoicing.json')

    store = oauth2client.file.Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        credentials = tools.run_flow(flow, store, flags)
        print('Storing credentials to ' + credential_path)
    return credentials


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
    fh = io.FileIO(file_name, mode='wb')
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
        print("Download %d%%." % int(status.progress() * 100))


def load_invoices(data_file_name):
    json_data = open(data_file_name).read()
    data = simplejson.loads(json_data)
    return data


def main():
    """The script load a file containing the invoices data and then print the one indicated.
    """

    # Parse argument
    parser = argparse.ArgumentParser(description='Process some integers.')
    parser.add_argument('data_file', metavar='F', type=str,
                       help='path to the data file')
    parser.add_argument('n_invoice', metavar='N', type=int,
                       help='invoice to print')
    args = parser.parse_args()

    # Load the inboice and select the one to process
    invoices = load_invoices(args.data_file)
    if not str(args.n_invoice) in invoices:
        print("Unknown invoice %i. Known invoices: %s" % (args.n_invoice, invoices.keys()))
        return
    invoice = invoices[str(args.n_invoice)]
    invoice['number'] = args.n_invoice

    # find the 'DriveInvoicing' directory and look for the file 'Template' inside it
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    drive_service = discovery.build('drive', 'v3', http=http)

    folder_id = get_folder(drive_service, 'DriveInvoicing')['id']
    template_id = get_content(drive_service, 'Template', folder_id)['id']

    # Copy the template
    invoice_doc_id = copy_file(drive_service, template_id, 'Invoice_%i' % invoice['number'], folder_id)['id']

    # Run the script to fill the template
    script_service = discovery.build('script', 'v1', http=http)
    paymentDays = None
    currency = DEFAULT_CURRENCY
    if 'currency' in invoice:
        currency = invoice['currency']
    if 'paymentDays' in invoice:
        paymentDays = invoice['paymentDays']
    request = {"function": "insertData", "devMode": True, "parameters": [
        invoice_doc_id, invoice['number'], invoice['date'], invoice['noVAT'], invoice['client'],
        invoice['lines'], currency, paymentDays]}
    response = script_service.scripts().run(body=request, scriptId=SCRIPT_ID).execute()
    print("Execution response: %s" % str(response))

    # Download the PDF file
    download_file_as_pdf(drive_service, invoice_doc_id, 'Invoice_%i.pdf' % invoice['number'])

if __name__ == '__main__':
    main()