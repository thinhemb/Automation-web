import time
import traceback
import win32com.client as win32
import os
from configs import logger, email_subject
from email_contents import *
from utils import read_file_txt, write_file_txt


def send_emails(to, subject, body):
    # Create a new instance of the Outlook application
    outlook = win32.Dispatch('Outlook.Application')

    # Create a new mail item
    mail = outlook.CreateItem(0x0)

    # Set up the email details
    mail.To = to
    mail.Subject = subject
    mail.Body = body
    # Uncomment the following line to add an attachment
    # mail.Attachments.Add('C:\\path\\to\\your\\file.txt')

    # Send the email
    mail.Send()

    print('Email sent successfully!')


def check_email_exist(email):
    outlook = win32.Dispatch('Outlook.Application')

    try:
        namespace = outlook.GetNamespace('MAPI')
        recipient = namespace.CreateRecipient(email)

        exchange_user = recipient.AddressEntry.GetExchangeUser()
        if exchange_user:
            return True
        else:
            return False
    except Exception as e:
        logger.warning(f"Exception in check_email_exist({email}).{e}\n{traceback.format_exc()}")
        return False


def extract_and_sending_email(sending_email_txt):
    # For Reset Password Email

    if not os.path.exists(sending_email_txt):
        logger.warning(f"Error while sending email to file {sending_email_txt} is not exist!")
        time.sleep(10)
        return

    # temp_file_path = os.path.splitext(sending_email_txt)[0] + "_tmp.txt"

    data_send_email = read_file_txt(sending_email_txt)
    write_file_txt(sending_email_txt, "", end=False)
    if len(data_send_email) < 1:
        logger.infor(f"File sending_email.txt no data!")
        return

    list_temp = []
    email = None
    for line in data_send_email:
        if line.strip() == '':
            continue
        try:
            data = line.split(' ')
            knt = data[0].strip()
            email = data[1].strip()
            code = data[2].strip()
            url = data[3].strip()
            contents = create_reset_pw_content(knt, url, code)

            if not check_email_exist(email=email):
                logger.warning(f"Email seem not exist {email}")
                continue

            send_emails(email, email_subject, contents)
            logger.infor(f"send_emails to '{email}' success!")
        except Exception as e:
            logger.warning(f"Error while sending email to {email}.{e}.\n{traceback.format_exc()}")
            list_temp.append(line)
    if len(list_temp) > 0:
        write_file_txt(sending_email_txt, list_temp, mode='w+')


def send_reminder_mail(json_data):
    # Opening JSON file
    for email in json_data['email']:
        if not check_email_exist(email):
            logger.warning(f"email '{email}' does not exist")
            continue

        contents, email_sub = create_reminder_content(json_data["dept"], json_data["link_tool"], json_data["link"])
        send_emails(email, email_sub, contents)
        logger.infor(f"send_reminder_mail() to '{email}' success!")


def send_request_sharepoint_mail(json_data):
    # Opening JSON file
    for email in json_data['email']:
        if not check_email_exist(email):
            logger.warning(f"email '{email}' does not exist")
            continue

        contents, email_sub = create_request_sharepoint_content(json_data['current_sharepoint'],
                                                                json_data['link_request_sharepoint'],
                                                                json_data['link_sharepoint_management'])
        send_emails(email, email_sub, contents)
        logger.infor(f"send_request_sharepoint_mail() to '{email}' success!")
