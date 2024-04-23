import logging
import os
import re
import uuid
from datetime import datetime, timedelta, date

import msal
import requests.exceptions
from O365 import Account, FileSystemTokenBackend
from keboola.component.base import ComponentBase
from keboola.component.exceptions import UserException
from keboola.utils import parse_datetime_interval

# configuration variables
KEY_SHAREPOINT = 'sharepoint'
KEY_O365 = 'o365'
KEY_CLIENT_ID = 'client_id'
KEY_CLIENT_SECRET = '#client_secret'
KEY_TENANT_ID = 'tenant_id'
KEY_USERNAME = 'username'
KEY_PASSWORD = '#password'
KEY_AUTHORITY = 'authority'
KEY_HOSTNAME = 'hostname'
KEY_URL = 'url'
KEY_FOLDER_SUFFIX = 'folder_suffix'

KEY_MAIN_FOLDER_PATH = 'main_folder_path'
KEY_OPERATION_TYPE = 'operation_type'
KEY_DATE_FROM = 'date_from'
KEY_DATE_TO = 'date_to'
KEY_FILTER_DAY = 'filter_day'

# list of mandatory parameters => if some is missing,
# component will fail with readable message on initialization.
REQUIRED_PARAMETERS = [KEY_SHAREPOINT, KEY_O365]


class Component(ComponentBase):

    def __init__(self):
        super().__init__()
        self.sharepoint_drive = None
        self.token_file_name = str(uuid.uuid4())
        self.scopes = ["Files.ReadWrite.All"]
        # set logging level
        logging.getLogger('O365.drive').setLevel(logging.CRITICAL)

    def run(self):

        self.validate_configuration_parameters(REQUIRED_PARAMETERS)
        params = self.configuration.parameters
        date_from = params[KEY_DATE_FROM]
        date_to = params[KEY_DATE_TO]
        operation_type = params[KEY_OPERATION_TYPE]
        main_folder_path = params[KEY_MAIN_FOLDER_PATH]
        sharepoint_params = params[KEY_SHAREPOINT]
        o365_params = params[KEY_O365]
        folder_suffix = params[KEY_FOLDER_SUFFIX]
        filter_day = params.get(KEY_FILTER_DAY)

        # create temp folder to store the token file in. The token name is random.
        self.create_temp_folder()
        self.get_token(sharepoint_params, o365_params)

        account = self.authenticate_o365_account(o365_params)
        self.sharepoint_drive = self.get_sharepoint_drive(account, o365_params)

        dt_format = '%Y-%m-%d'
        try:
            start_date, end_date = parse_datetime_interval(date_from, date_to, dt_format)
        except TypeError:
            raise UserException(f"Unsupported date strings: {date_from}, {date_to}")
        except ValueError as e:
            raise e

        start_date = self.get_datetime(start_date)
        end_date = self.get_datetime(end_date)

        days_to_process = self.get_dates_between(start_date, end_date)
        for day in days_to_process:
            logging.info(f"Processing date: {day}")
            self.process_files(day, operation_type, main_folder_path, folder_suffix, filter_day)

    @staticmethod
    def get_dates_between(start_date, end_date):
        dates = []
        current_date = start_date
        while current_date < end_date:
            date_str = current_date.strftime('%Y-%m-%d')
            dates.append(date_str)
            current_date += timedelta(days=1)
        return dates

    @staticmethod
    def get_datetime(date_str):
        year, month, day = map(int, date_str.split("-"))
        return date(year, month, day)

    def get_input_files(self):
        files = self.get_input_files_definitions(only_latest_files=True)
        return files

    @staticmethod
    def subtract_one_day(date_string):
        date_obj = datetime.strptime(date_string, '%Y-%m-%d')
        new_date = date_obj - timedelta(days=1)
        return new_date.strftime('%Y_%m_%d')

    @staticmethod
    def extract_date(string):
        pattern = r'\d{4}_\d{2}_\d{2}'  # regex pattern to match date format yyyy_mm_dd
        match = re.search(pattern, string)
        if match:
            return match.group()
        else:
            return None

    def process_files(self, date_of_processing, operation_type, main_folder_path, folder_suffix, filter_day):

        folder = main_folder_path + date_of_processing
        if not folder.startswith("/"):
            folder = "/" + folder

        if folder_suffix:
            folder = folder + folder_suffix

        if operation_type == "upload":
            files = self.get_input_files()
            logging.debug(f"Found {len(files)}")
            for file in files:
                if filter_day:
                    filename_date = self.extract_date(file.name)
                    if self.subtract_one_day(date_of_processing) == filename_date:
                        self.upload(folder_name=folder, file=file)

                    self.upload(folder_name=folder, file=file)
                else:
                    self.upload(folder_name=folder, file=file)

        elif operation_type == "download":
            self.download(folder_name=folder)
        else:
            raise UserException(f"Invalid operation type: {operation_type}")

    def upload(self, folder_name, file):
        """Uploads file_name Sharepoint folder_name (with '/')"""
        try:
            folder = self.sharepoint_drive.get_item_by_path(folder_name)
        except requests.exceptions.HTTPError:
            logging.info(f"Folder {folder_name} does not exist. The component will attempt to create it.")
            logging.info(f"Trying to create new folder: {folder_name}")
            folder = self.sharepoint_drive.get_root_folder()
            self.create_new_onedrive_folder(folder, folder_name)
            folder = self.sharepoint_drive.get_item_by_path(folder_name)

        input_file_path = os.path.join(file.full_path)
        logging.info(f"Uploading file: {file.name}")

        folder.upload_file(item=input_file_path)

    def create_new_onedrive_folder(self, folder, path):
        path = path[1:]
        path_list = path.split("/")  # Split the string into a list using "/"
        current_path = ""
        for item in path_list:
            try:
                folder.create_child_folder(item)
                logging.info(f"Subfolder {item} created.")
            except requests.exceptions.HTTPError:
                logging.info(f"Subfolder {item} already exists.")
            finally:
                current_path += f"/{item}"
                folder = self.sharepoint_drive.get_item_by_path(current_path)

    def download(self, folder_name):
        """Downloads all files in a specified OneDrive folder."""
        try:
            onedrive_folder = self.sharepoint_drive.get_item_by_path(folder_name)
        except requests.exceptions.HTTPError:
            logging.warning(f"Folder {folder_name} not found on server.")
            onedrive_folder = None

        if onedrive_folder:
            for f in onedrive_folder.get_items():
                if f.is_file:
                    file_path = os.path.join(folder_name, f.name)
                    logging.info(f"Downloading file: {f.name}")
                    file = self.sharepoint_drive.get_item_by_path(file_path)
                    file.download(to_path=self.files_out_path)

                    file_def = self.create_out_file_definition(name=f.name, tags=["chatbot_analytics",
                                                                                  f"source_path: {file_path}"])
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
