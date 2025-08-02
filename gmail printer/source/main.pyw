'''
Author: Kyle Sherman
Created: 10/04/2024
Updated: 01/30/2025
'''

import configparser
import os
import google.auth
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import re
import base64
import time
from googleapiclient.errors import HttpError
import subprocess
import logging
import traceback
import sys
from datetime import datetime
from pathlib import Path
import win32api
import win32print

SCOPES = ['https://www.googleapis.com/auth/gmail.readonly', 'https://www.googleapis.com/auth/gmail.modify']

project_directory = None
source_directory = None
orders_directory = None

acrobat_path = None

credentials_file = None
token_file = None

scan_interval = None
sender_email = None

def log_and_print(message, level):
    # print(message)

    if level == "debug":
        logging.debug(message)
    elif level == "error":
        logging.error(message)
    elif level == "warning":
        logging.warning(message)
    elif level == "critical":
        logging.critical(message)
    else:
        logging.info(message)

# setup logging to a file
def setup_logging(directory):

    path = os.path.join(directory, 'logs')

    if not os.path.exists(path):
        os.makedirs(path)

    logfile = datetime.now().strftime(f"{path}/log_%Y-%m-%d_%H-%M-%S.log")

    #create a logger
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG) # set logger to the lowest level to catch everything

    # create file handler for logging to a file (all levels)
    file_handler = logging.FileHandler(logfile)
    file_handler.setLevel(logging.DEBUG)
    file_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    file_handler.setFormatter(file_formatter)

    # create a console handler for logging to the console (only info and above)
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    console_handler.setFormatter(console_formatter)

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    logging.info(f"Logging setup complete. Log file: {logfile}")

def log_uncaught_exceptions(exception_type, exception_value, exception_traceback):
    if issubclass(exception_type, KeyboardInterrupt):
        sys.__excepthook__(exception_type, exception_value, exception_traceback)
        return
    logging.critical("Unhandled exception", exc_info=(exception_type, exception_value, exception_traceback))       

# set the glocal exception hook to log uncaught exceptions
sys.excepthook = log_uncaught_exceptions

def load_config(file_path):
    current_directory = os.getcwd()

    # file_path = os.path.join(current_directory, file_path)
    config = configparser.ConfigParser()
    config.read(file_path)

    # print(file_path)
    return config

def print_pdf_with_acrobat(file_path, order_number):
    # Print pdf using adobe reader via subprocess
    try:
        # subprocess.run([acrobat_path, '/h', '/t', file_path], check=True)
        # print(f"printing order: {order_number}")
        printer_name = win32print.GetDefaultPrinter()
        win32api.ShellExecute(0, "print", file_path, f'/d:"{printer_name}"', ".", 0)
        log_and_print(f"printing order: {order_number}", "info")
    except subprocess.CalledProcessError as e:
        # print(f"failed to print {file_path}. Error: {e}")
        log_and_print(f"failed to print {file_path}. Error: {e}", "error")
    except FileNotFoundError:
        # print(f"Adoby reader not found at {acrobat_path}. Please ensure it is set to the correct location")
        log_and_print(f"Adoby reader not found at {acrobat_path}. Please ensure it is set to the correct location", "error")


def save_attachements(service, message, order_number):
    for part in message['payload']['parts']:
        if part['filename'] and part['filename'].lower().endswith('.pdf'):
            attachment_id = part['body'].get('attachmentId')
            if attachment_id:
                attachment = service.users().messages().attachments().get(
                    userId='me', messageId=message['id'], id=attachment_id
                ).execute()

                data = base64.urlsafe_b64decode(attachment['data'].encode('UTF-8'))

                # save the file
                file_path = os.path.join(orders_directory, part['filename'])
                with open(file_path, 'wb') as f:
                    f.write(data)

                # try to print the attachment
                try:
                    print_pdf_with_acrobat(file_path, order_number)
                except Exception as e:
                    log_and_print(f"failed to print order: {order_number}. Error {e}", "error")

def check_emails(service, sender_email):

    retries = 0
    backoff = 2
    max_retries = 5

    ''' check for new emails from a specific sender '''
    # search query to filter emails from a specific sender
    email_query = f'from:{sender_email} subject:"INCOMING DELIVERY ORDER #" is:unread'

    while retries < 5:
        service = gmail_service()
        
        try:
            
            # fetch messages
            results = service.users().messages().list(userId='me', q=email_query, maxResults=1).execute()
            messages = results.get('messages', [])

            if not messages:
                return False
            else:
                for message in messages:
                    email = service.users().messages().get(userId='me', id=message['id']).execute()
                    for header in email['payload']['headers']:
                        if header['name'] == 'Subject':
                            subject = header['value']

                            order_number = re.search(r'#(\d+)', subject).group(1)

                            log_and_print(f"New Dine In Order: {order_number}", "info")

                            if 'parts' in email['payload']:
                                save_attachements(service, email, order_number)

                            # mark the email as read
                            service.users().messages().modify(
                                userId='me', id=message['id'], body={'removeLabelIds': ['UNREAD']}
                            ).execute()

            break # break out of the loop if it fails
        
        except HttpError as error:
            log_and_print(f"an error has occured: {error}", "error")

            if error.resp.status in [429, 500, 503]: # check for retryable errors
                retries += 1
                delay = backoff * (2 ** retries) # exponential backoff
                log_and_print(f"retrying after {delay} seconds... (attempt {retries} / {max_retries})", "warning")
                time.sleep(delay)
            else:
                print(f"a non retriable error has occured. Exting {error}", "critical")
                break # non retryable error. Exit

        except ConnectionError as ce:
            # handle connection errors (i.e. no internet connection)
            log_and_print(f"connection error occurred: {ce}", "error")
            retries += 1
            delay = backoff * (2 ** retries) # exponential backoff
            log_and_print(f"retrying after {delay} seconds... (attempt {retries} / {max_retries})", "warning")
            time.sleep(delay)

        except Exception as e:
            # catch any other exceptions
            log_and_print(f"an unexpected error has occurred: {e}", "error")
            break

    if retries == max_retries:
        log_and_print("failed to connect after multiple attempts. Exiting.", "critical")
    

def get_local_path():

    if getattr(sys, 'frozen', False):
        current_path = os.path.dirname(sys.executable)
    else:
        current_path = os.path.dirname(os.path.abspath(__file__))

    parent_path = os.path.dirname(current_path)

    return parent_path

def load_credentials():
    credentials = None
    
    # load the credentials, if any
    if os.path.exists(token_file):
        credentials = Credentials.from_authorized_user_file(token_file, SCOPES)

    return credentials

def save_credentials(credentials):
    # save the credentials for future use
    with open(token_file, 'w') as token:
        token.write(credentials.to_json())
        
    
def refresh_credentials(credentials):
    creds = credentials
    
    # refresh the token if needed
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
        save_credentials(creds)
    else:
        flow = InstalledAppFlow.from_client_secrets_file(credentials_file, SCOPES)
        creds = flow.run_local_server(port=0, access_type="offline", prompt="consent")
        save_credentials(creds)
    
    return creds
    
def get_valid_credentials():
    creds = load_credentials()
    
    if not creds or not creds.valid:
        creds = refresh_credentials(creds)
        
    return creds
    
    
def gmail_service():
    creds = get_valid_credentials()
    return build('gmail', 'v1', cache_discovery=False, credentials=creds)
    

def main():
    '''shows basic usage of gmail api
        lists the subject of the last 3 messages in the user's inbox'''
    
    creds = None
    
    creds = get_valid_credentials()

    last_print_time = time.time()
    no_message_print_interval = 30

    try:
        while True:
            service = gmail_service()
            
            new_email = check_emails(service, sender_email)

            current_time = time.time()

            if not new_email:
                if current_time - last_print_time >= no_message_print_interval:
                    log_and_print("No new orders in last 30 seconds", "info")
                    last_print_time = current_time

            time.sleep(no_message_print_interval) # check email every 10 seconds
    
    except KeyboardInterrupt:
        log_and_print("program terminated by the user.", "info")
    
    except Exception as e:
        log_and_print(f"Application has crashed: {e}", "critical")
        log_and_print(f"{traceback.format_exc()}", "debug")

if __name__ == "__main__":

    current_dir = get_local_path()

    setup_logging(current_dir)
    
    sys.stdout.reconfigure(line_buffering=True)

    try:

        config_dir = os.path.join(current_dir, 'config')

        config_path = os.path.join(config_dir, "config.config")

        if not os.path.exists(config_path):
            raise FileNotFoundError(f"Config file not found at {config_path}")

        config = load_config(config_path)

        log_and_print(f"config loaded at {config_path}", "info")

    except FileNotFoundError as error:
        log_and_print(error, "critical")
        sys.exit(1)
    
    except Exception as error:
        log_and_print(error, "critical")
        sys.exit(1)

    project_directory = config['files']['project_directory']

    source_directory = config['files']['source_directory']

    orders_directory = config['files']['orders_directory']

    config_directory = config['files']['config_directory']

    credentials_file = os.path.join(config_directory, config['files']['credentials_file'])
    token_file = os.path.join(config_directory, config['files']['pickle_file'])
    
    scan_interval = int(config['settings']['scan_interval'])

    sender_email = config['settings']['sender_email']

    acrobat_path = config['settings']['acrobat_path']

    log_and_print(f"Starting Dine-In order scaner - by Kyle Sherman", "info")
    log_and_print(f"Checking for emails from {sender_email}", "info")

    main()