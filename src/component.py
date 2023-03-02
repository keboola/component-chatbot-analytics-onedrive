import logging
import os
import uuid

import msal
from O365 import Account, FileSystemTokenBackend

from keboola.component.base import ComponentBase
from keboola.component.exceptions import UserException
from keboola.utils import parse_datetime_interval

# configuration variables
KEY_SHAREPOINT = 'sharepoint'
KEY_CLIENT_ID = 'client_id'
KEY_CLIENT_SECRET = '#client_secret'
KEY_TENANT_ID = 'tenant_id'
KEY_USERNAME = 'username'
KEY_PASSWORD = '#password'
KEY_AUTHORITY = 'authority'
KEY_HOSTNAME = 'hostname'
KEY_URL = 'url'

KEY_MAIN_FOLDER_PATH = 'main_folder_path'
KEY_OPERATION_TYPE = 'operation_type'
KEY_FILENAME_PREFIX = 'filename_prefix'
KEY_DATE_OF_PROCESSING = 'date_of_processing'

# list of mandatory parameters => if some is missing,
# component will fail with readable message on initialization.
REQUIRED_PARAMETERS = [KEY_SHAREPOINT]


class Component(ComponentBase):

    def __init__(self):
        super().__init__()
        self.sharepoint_drive = None
        self.token_file_name = str(uuid.uuid4())
        self.scopes = ["Files.ReadWrite.All", "offline_access"]

    def run(self):

        self.validate_configuration_parameters(REQUIRED_PARAMETERS)
        params = self.configuration.parameters

        sharepoint_params = params["sharepoint"]
        o365_params = params["o365"]

        self.get_token(sharepoint_params)

        # create temp folder to store the token file in. The token name is random.
        self.create_temp_folder()

        account = self.authenticate_o365_account(o365_params)
        self.sharepoint_drive = self.get_sharepoint_drive(account, o365_params)

        self.process_files(main_folder_path=os.path.join(self.files_in_path, params["main_folder_path"]),
                           date=params[KEY_DATE_OF_PROCESSING],
                           filename_prefix=params[KEY_FILENAME_PREFIX],
                           operation_type=params[KEY_OPERATION_TYPE],
                           params=params)

    @staticmethod
    def get_date_of_processing(date):
        dt_str_1 = date
        dt_str_2 = "today"
        dt_format = "%Y-%m-%d"
        date_of_processing, _ = parse_datetime_interval(dt_str_1, dt_str_2, dt_format)
        return str(date_of_processing)

    @staticmethod
    def list_files_with_prefix(folder_path, prefix):
        files = os.listdir(folder_path)
        return [f for f in files if f.startswith(prefix)]

    def process_files(self, main_folder_path, date, filename_prefix, operation_type, params):
        date_of_processing = self.get_date_of_processing(date)
        subfolder_path = os.path.join(main_folder_path, date_of_processing)
        files = self.list_files_with_prefix(subfolder_path, filename_prefix)
        folder = params[KEY_MAIN_FOLDER_PATH] + date_of_processing
        if not folder.startswith("/"):
            folder = "/" + folder
        for file in files:
            if operation_type == "upload":
                self.upload(folder_name=folder, file_name=file)
            elif operation_type == "download":
                self.download(folder_name=folder, file_name=file, to_path=self.files_out_path)
            else:
                raise UserException(f"Invalid operation type: {operation_type}")

    def upload(self, folder_name, file_name):
        """Uploads file_name Sharepoint folder_name (with '/')"""
        folder = self.sharepoint_drive.get_item_by_path('/' + folder_name)
        folder.upload_file(item=file_name)

    def download(self, folder_name, file_name, to_path):
        """Downloads file_name from Sharepoint folder_name (with "/") to to_path folder"""
        folder_name = folder_name + file_name
        file = self.sharepoint_drive.get_item_by_path('/' + folder_name)
        file.download(to_path=to_path)

    @staticmethod
    def get_sharepoint_drive(account, o365_params):
        sharepoint = account.sharepoint()
        site = sharepoint.get_site(o365_params[KEY_HOSTNAME], o365_params[KEY_URL])
        return site.get_default_document_library()

    def authenticate_o365_account(self, o365_params):
        credentials = (o365_params[KEY_CLIENT_ID], o365_params[KEY_CLIENT_SECRET])
        temp_folder = os.path.join(self.data_folder_path, "temp")
        token_backend = FileSystemTokenBackend(token_path=temp_folder,
                                               token_filename=self.token_file_name)

        account = Account(credentials, tenant_id=o365_params[KEY_TENANT_ID], token_backend=token_backend)
        if not account.is_authenticated:
            raise UserException("Cannot Authenticate o365 account.")
        return account

    def create_temp_folder(self):
        temp_folder_path = os.path.join(self.data_folder_path, "temp")
        if not os.path.exists(temp_folder_path):
            os.makedirs(temp_folder_path)

    def get_token(self, sharepoint_params):
        """Retrieves and saves the token to temp folder with random filename generated with uuid."""

        # Create a preferably long-lived app instance which maintains a token cache.
        app = msal.PublicClientApplication(sharepoint_params[KEY_CLIENT_ID],
                                           authority=sharepoint_params[KEY_AUTHORITY])

        result = None

        # Firstly, check the cache to see if this end user has signed in before
        accounts = app.get_accounts(username=sharepoint_params[KEY_USERNAME])
        if accounts:
            logging.info("Account(s) exists in cache, probably with token too. Let's try.")
            result = app.acquire_token_silent(sharepoint_params[self.scopes], account=accounts[0])

        if not result:
            logging.info("No suitable token exists in cache. Let's get a new one from AAD.")
            # See this page for constraints of Username Password Flow.
            # https://github.com/AzureAD/microsoft-authentication-library-for-python/wiki/Username-Password-Authentication
            result = app.acquire_token_by_username_password(
                sharepoint_params[KEY_USERNAME], sharepoint_params[KEY_PASSWORD],
                scopes=sharepoint_params[self.scopes])

        token = str(result).replace('\'', '"')
        if not token:
            raise UserException("Cannot retrieve token.")

        with open(self.token_file_name, 'w') as f:
            f.write(token)


"""
        Main entrypoint
"""
if __name__ == "__main__":
    try:
        comp = Component()
        # this triggers the run method by default and is controlled by the configuration.action parameter
        comp.execute_action()
    except UserException as exc:
        logging.exception(exc)
        exit(1)
    except Exception as exc:
        logging.exception(exc)
        exit(2)
