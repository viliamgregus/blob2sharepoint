#import libs
from modules.sharepoint import SharePoint
import os
from azure.storage.blob import BlobClient
from flask import request
from shareplum import Site, Office365
from shareplum.site import Version

#azure storage
#user input
STORAGE_ACCOUNT_NAME=request.args.get('storage_account_name')
CONTAINER_NAME=request.args.get('container_name')
FILE_NAME = request.args.get('file_name')
#derived
azure_access_key=""
azure_conn_str=f'DefaultEndpointsProtocol=https;AccountName={STORAGE_ACCOUNT_NAME};AccountKey={azure_access_key};EndpointSuffix=core.windows.net'
blob = BlobClient.from_connection_string(
        conn_str=azure_conn_str,
        container_name=CONTAINER_NAME,
        blob_name=FILE_NAME,
        credential=azure_access_key
    )

#staging
#user input
STAGING_FILE = request.args.get('staging_file')
#derived
file_dir_path = f".\{STAGING_FILE}"

#sharepoint
#user input
USERNAME = request.args.get('user')
PASSWORD = request.args.get('password')
SHAREPOINT_URL = request.args.get('url')
SHAREPOINT_SITE = request.args.get('site')
FOLDER_NAME = request.args.get('folder_path')
SHAREPOINT_DOC = request.args.get('doc_library')

class SharePoint:
    def auth(self):
        self.authcookie = Office365(SHAREPOINT_URL, username=USERNAME, password=PASSWORD).GetCookies()
        self.site = Site(SHAREPOINT_SITE, version=Version.v365, authcookie=self.authcookie)

        return self.site

    def connect_folder(self, FOLDER_NAME):
        self.auth_site = self.auth()

        self.sharepoint_dir = '/'.join([SHAREPOINT_DOC, FOLDER_NAME])
        self.folder = self.auth_site.Folder(self.sharepoint_dir)

        return self.folder

    def upload_file(self, file, FILE_NAME, FOLDER_NAME):
        self._folder = self.connect_folder(FOLDER_NAME)

        with open(file, mode='rb') as file_obj:
            file_content = file_obj.read()

        self._folder.upload_file(file_content, FILE_NAME)

    def delete_file(self, FILE_NAME, FOLDER_NAME):

        self._folder = self.connect_folder(FOLDER_NAME)

        self._folder.delete_file(FILE_NAME)


def from_azure_to_sharepoint():
    # Define blob file
    blob = BlobClient.from_connection_string(
            conn_str=azure_conn_str,
            container_name=CONTAINER_NAME,
            blob_name=FILE_NAME,
            credential=azure_access_key
    )
    # Download blob to staging file in directory of the py file
    with open(file_dir_path, "wb") as my_blob:
        download_stream = blob.download_blob()
        my_blob.write(download_stream.readall())

    # Upload file from staging directory to Sharepoint
    SharePoint().upload_file(file_dir_path, FILE_NAME, FOLDER_NAME)

    # remove the staging file
    os.remove(file_dir_path)