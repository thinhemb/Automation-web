import configparser
import os
from my_log import OperationLog
from datetime import datetime
import traceback
import ctypes

logger = OperationLog(log_name=f"Log [{datetime.now().strftime('%Y-%m-%d')}]")

WORKING_DIRECTORY = os.path.dirname(__file__)

RESOURCES_PATH = os.path.join(WORKING_DIRECTORY, "Resources")

CONFIGS_TXT = os.path.join(RESOURCES_PATH, "configs.ini")

PROJECT_NAME = "EIW_Upload_Support"

PROJECT_REMINDER_EXE_PATH = os.path.join(WORKING_DIRECTORY, PROJECT_NAME + ".exe")

waiting_5_minute = 300

try:
    configs_ini = configparser.ConfigParser()
    configs_ini.read(CONFIGS_TXT)

    scan_time = configs_ini['Settings']['scan_time']
    sharepoint_noti_time = configs_ini['Settings']['sharepoint_noti_time']
    sharepoint_noti_mess = configs_ini['Settings']['sharepoint_noti_mess']

    email_scan_time = configs_ini['Settings']['email_scan_time']
    email_subject = configs_ini['Settings']['email_subject']

    input_folder = configs_ini['Settings']['input_folder']
    # sharepoint_output_folder = configs_ini['Settings']['sharepoint_output_folder']
    # natv_output_folder = configs_ini['Settings']['natv_output_folder']
    # nat_output_folder = configs_ini['Settings']['nat_output_folder']
    # nml_output_folder = configs_ini['Settings']['nml_output_folder']

    server_username = configs_ini['Settings']['server_username']
    server_password = configs_ini['Settings']['server_password']

    # move_files_txt = os.path.join(input_folder, configs_ini['Settings']['move_files_txt'])
    # delete_files_txt = os.path.join(input_folder, configs_ini['Settings']['delete_files_txt'])
    # notify_txt = os.path.join(input_folder, configs_ini['Settings']['notify_txt'])
    sending_email_txt = os.path.join(input_folder, configs_ini['Settings']['sending_email_txt'])

    remind_mail_day = list(configs_ini['Settings']['remind_mail_day'])
    delete_tool_web_day = int(configs_ini['Settings']['delete_tool_web_day'])

    mail_admin = configs_ini['Settings']['mail_admin'].split(', ')
    link_request_sharepoint = configs_ini['Settings']['link_request_sharepoint']
    link_sharepoint_management = configs_ini['Settings']['link_sharepoint_management']
    check_empty_sharepoint_json = os.path.join(input_folder, configs_ini['Settings']['check_empty_sharepoint_json'])

    delete_tool_web = os.path.join(input_folder, configs_ini['Settings']['delete_tool_web'])
    check_existed_sharepoint_json = os.path.join(input_folder, configs_ini['Settings']['check_existed_sharepoint_json'])

    delete_tool_sharepoint_json = os.path.join(input_folder, configs_ini['Settings']['delete_tool_sharepoint_json'])
    delete_unuse_folder_json = os.path.join(input_folder, configs_ini['Settings']['delete_unuse_folder_json'])

    move_folder_json = os.path.join(input_folder, configs_ini['Settings']['move_folder_json'])

    count_storage_sharepoint_json = os.path.join(input_folder, configs_ini['Settings']['count_storage_sharepoint_json'])
    reach_40GB_sharepoint_txt = os.path.join(input_folder, configs_ini['Settings']['reach_40GB_sharepoint_txt'])
    reach_45GB_sharepoint_txt = os.path.join(input_folder, configs_ini['Settings']['reach_45GB_sharepoint_txt'])

    new_sharepoint_json = os.path.join(input_folder, configs_ini['Settings']['new_sharepoint_json'])
    permission_group_json = os.path.join(input_folder, configs_ini['Settings']['permission_group_json'])
    flag_folder = configs_ini['Settings']['flag'].split(',')

    # Convert string to time object
    scan_time = datetime.strptime(scan_time, "%H:%M:%S").time()
    sharepoint_noti_time = datetime.strptime(sharepoint_noti_time, "%H:%M:%S").time()
    email_scan_time = datetime.strptime(email_scan_time, "%H:%M:%S").time()
    # Convert time object to total seconds
    scan_time = scan_time.hour * 3600 + scan_time.minute * 60 + scan_time.second
    sharepoint_noti_time = (sharepoint_noti_time.hour * 3600 + sharepoint_noti_time.minute * 60
                            + sharepoint_noti_time.second)
    email_scan_time = email_scan_time.hour * 3600 + email_scan_time.minute * 60 + email_scan_time.second

    reach_40GB_sharepoint_capacity = int(configs_ini['Settings']['reach_40GB_sharepoint_capacity'])
    reach_45GB_sharepoint_capacity = int(configs_ini['Settings']['reach_45GB_sharepoint_capacity'])

    time_sleep_check_empty_sharepoint = int(configs_ini['Settings']['time_sleep_check_empty_sharepoint'])
    time_sleep_check_existed_sharepoint = int(configs_ini['Settings']['time_sleep_check_existed_sharepoint'])
    time_sleep_delete_tool_sharepoint = int(configs_ini['Settings']['time_sleep_delete_tool_sharepoint'])
    time_sleep_delete_unuse_folder = int(configs_ini['Settings']['time_sleep_delete_unuse_folder'])
    time_sleep_move_folder = int(configs_ini['Settings']['time_sleep_move_folder'])
    time_sleep_count_storage_sharepoint = int(configs_ini['Settings']['time_sleep_count_storage_sharepoint'])
    time_sleep_new_sharepoint = int(configs_ini['Settings']['time_sleep_new_sharepoint'])
    time_sleep_new_sharepoint_and_permission = int(configs_ini['Settings']['time_sleep_new_sharepoint_and_permission'])
    time_sleep_permission_group = int(configs_ini['Settings']['time_sleep_permission_group'])

    email_NAT = configs_ini['Settings']['email_NAT']
    email_NATV = configs_ini['Settings']['email_NATV']
    export_user_email_csv = str(configs_ini['Settings']['export_user_email_csv'])

    start_time_one_slot = str(configs_ini['Settings']['start_time_one_slot'])
    end_time_n_slot = str(configs_ini['Settings']['end_time_n_slot'])

    output_check_user_outlook_csv = os.path.join(input_folder, configs_ini['Settings']['output_check_user_outlook_csv'])
    input_check_user_outlook_csv = os.path.join(input_folder, configs_ini['Settings']['input_check_user_outlook_csv'])
    config_check_user_outlook_csv = os.path.join(input_folder, configs_ini['Settings']['config_check_user_outlook_csv'])
    update_current_user_txt = os.path.join(input_folder, configs_ini['Settings']['update_current_user_txt'])

except Exception as ex:
    print(ex)
    logger.error('Error while Parsing configs.ini')
    logger.error(traceback.format_exc())
    # Display an error message box # 0x10 is the value for the OK button
    ctypes.windll.user32.MessageBoxW(0, "Error while Parsing configs.ini", "EIW Upload Support Error", 0x10)
    exit()
