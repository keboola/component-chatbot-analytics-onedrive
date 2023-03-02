import logging
import os
import uuid
from datetime import datetime, timedelta

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
KEY_DATE_OF_PROCESSING = 'date_of_processing'

# list of mandatory parameters => if some is missing,
# component will fail with readable message on initialization.
REQUIRED_PARAMETERS = [KEY_SHAREPOINT]


class Component(ComponentBase):

    def __init__(self):
        super().__init__()
        self.sharepoint_drive = None
        self.token_file_name = str(uuid.uuid4())
        self.scopes = ["Files.ReadWrite.All"]

    def run(self):

        self.validate_configuration_parameters(REQUIRED_PARAMETERS)
        params = self.configuration.parameters

        sharepoint_params = params["sharepoint"]
        o365_params = params["o365"]

        self.process_files(main_folder_path=os.path.join(self.files_in_path, params["main_folder_path"]),
                           date=params[KEY_DATE_OF_PROCESSING],
                           operation_type=params[KEY_OPERATION_TYPE],
                           params=params)

        self.get_token(sharepoint_params, o365_params)

        # create temp folder to store the token file in. The token name is random.
        self.create_temp_folder()

        account = self.authenticate_o365_account(o365_params)
        self.sharepoint_drive = self.get_sharepoint_drive(account, o365_params)

    @staticmethod
    def get_date_of_processing(date):
        dt_str_1 = date
        dt_str_2 = "today"
        dt_format = "%Y-%m-%d"
        date_of_processing, _ = parse_datetime_interval(dt_str_1, dt_str_2, dt_format)
        return str(date_of_processing)

    def get_input_files(self):
        files = self.get_input_file_definitions_grouped_by_tag_group(only_latest_files=True)
        return files

    @staticmethod
    def subtract_one_day(date_string):
        date_obj = datetime.strptime(date_string, '%Y-%m-%d')
        new_date = date_obj - timedelta(days=1)
        return new_date.strftime('%Y-%m-%d')

    def process_files(self, main_folder_path, date, operation_type, params):
        date_of_processing = self.get_date_of_processing(date)
        logging.info(f"Processing date: {date_of_processing}")

        folder = params[KEY_MAIN_FOLDER_PATH] + date_of_processing
        if not folder.startswith("/"):
            folder = "/" + folder

        if operation_type == "upload":
            files = self.get_input_files()
            print(files)
            exit()
            for file in files:
                self.upload(folder_name=folder, file_name=file)
        elif operation_type == "download":
            self.download(folder_name=folder)
        else:
            raise UserException(f"Invalid operation type: {operation_type}")

    def upload(self, folder_name, file_name):
        """Uploads file_name Sharepoint folder_name (with '/')"""
        folder = self.sharepoint_drive.get_item_by_path(folder_name)
        input_file_path = os.path.join(self.files_in_path, file_name)
        logging.info(f"Uploading file: {file_name}")

        # Create OneDrive folder if it does not exist
        try:
            logging.info(f"Trying to create new folder: {folder_name}")
            folder.create_child_folder(folder_name)
        except Exception as e:
            logging.error(f"Cannot create new folder {folder_name}, exception: {e}")

        folder.upload_file(item=input_file_path)

    def download(self, folder_name):
        """Downloads all files in a specified OneDrive folder."""
        onedrive_folder = self.sharepoint_drive.get_item_by_path(folder_name)
        for f in onedrive_folder.get_items():
            file_path = os.path.join(folder_name, f)
            logging.info(f"Downloading file: {f}")
            file = self.sharepoint_drive.get_item_by_path(file_path)
            file.download(to_path=self.files_out_path)

            file_def = self.create_out_file_definition(name=f, tags=["chatbot_analytics"])
            self.write_manifest(file_def)

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

    def get_token(self, sharepoint_params, o365_params):
        """Retrieves and saves the token to temp folder with random filename generated with uuid."""

        # Create a preferably long-lived app instance which maintains a token cache.
        app = msal.PublicClientApplication(o365_params[KEY_CLIENT_ID],
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
                scopes=self.scopes)

        token = str(result).replace('\'', '"')
        if not token:
            raise UserException("Cannot retrieve token.")

        temp_path = os.path.join(self.data_folder_path, "temp")
        token_path = os.path.join(temp_path, self.token_file_name)
        with open(token_path, 'w') as f:
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
