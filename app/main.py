import sys
import subprocess
import threading
import pythoncom
from playwright.sync_api import sync_playwright

from start_with_computer import set_startup
from send_emails import *
from utils import *
from configs import *
from task_sharepoint import TaskSharepoint
from auto_sharepoint import AutoSharepoint
from check_user_outlook import UserOutlook


def notify_setup_success():
    # Determine the directory of the executable
    if getattr(sys, 'frozen', False):
        # If the application is run as a bundle, the PyInstaller bootloader
        # extends the sys module by a flag frozen=True and sets the app
        # path into variable _MEIPASS'.
        exe_dir = os.path.dirname(sys.executable)
    else:
        exe_dir = os.path.dirname(os.path.abspath(__file__))
    # print(exe_dir)

    # Create a file in the same directory
    file_path = os.path.join(exe_dir, 'Setup_Successfully.txt')
    if os.path.isfile(file_path):
        pass
    else:
        # Open the file in write mode and write some content
        content = f'{PROJECT_NAME} Set Up Successfully!!!\n{datetime.now()}'
        write_file_txt(file_path, content, mode='a', end=False)

        ctypes.windll.user32.MessageBoxW(0, f"'{PROJECT_NAME}' Tool SetUp Successfully", "SetUp", 1)


# Function to mount network drive
def mount_network_drive():
    try:
        # Replace USERNAME and PASSWORD with your actual credentials
        username = server_username
        password = server_password
        network_path = input_folder

        # Command to mount the network drive
        command = f'net use {network_path} /user:{username} {password}'

        # Execute the command
        subprocess.run(command, shell=True)
    except Exception as e:
        logger.error(f'Error while mount_network_drive({input_folder}) : {e}')
        logger.error(traceback.format_exc())
        # Display an error message box # 0x10 is the value for the OK button
        ctypes.windll.user32.MessageBoxW(0, "Error while mount_network_drive()", "EIW Upload Support Error", 0x10)
        exit()


def is_within_runtime_range(start="04:55:00", end="19:00:00"):
    """Check if the current time is between 5 AM and 7 PM."""
    current_time = datetime.now().time()
    start_time = datetime.strptime(start, "%H:%M:%S").time()
    end_time = datetime.strptime(end, "%H:%M:%S").time()
    return start_time <= current_time <= end_time


def delete_unuse_folder_loop():
    show_log_no_data = True
    while True:
        try:
            if not is_within_runtime_range(start=start_time_one_slot, end=set_end_time(start_time_one_slot)):
                break
            start = time.perf_counter()
            data = read_json(delete_unuse_folder_json)
            if not len(data):
                if show_log_no_data:
                    logger.infor(f"File {delete_unuse_folder_json} no data!")
                show_log_no_data = False
                time.sleep(30)
                continue
            show_log_no_data = True
            logger.infor(f"File '{delete_unuse_folder_json}' have data!")
            task = TaskSharepoint("delete_unuse_folder_loop")

            for item in data:
                task.delete_folder(item)

            logger.infor(f"Function delete_unuse_folder_loop is Successfully!")

            if (time.perf_counter() - start) < time_sleep_delete_unuse_folder - 1:
                time.sleep(time_sleep_delete_unuse_folder - (time.perf_counter() - start))
            # run 1 time in 1 day

        except Exception as e:
            logger.error(f"Error in delete_unuse_folder_loop(): {e}")
            logger.error(traceback.format_exc())


def check_empty_sharepoint_loop():
    pythoncom.CoInitialize()
    show_log_no_data = True
    while True:
        try:
            if not is_within_runtime_range(start=start_time_one_slot, end=set_end_time(start_time_one_slot)):
                break
            start = time.perf_counter()
            data = read_json(check_empty_sharepoint_json)
            if not len(data):
                if show_log_no_data:
                    logger.infor(f"File '{check_empty_sharepoint_json}' no data!")
                show_log_no_data = False
                time.sleep(15)
                continue
            show_log_no_data = True
            logger.infor(f"File '{check_empty_sharepoint_json}' have data!")
            task = TaskSharepoint("check_empty_sharepoint_loop")
            list_folder = []
            data_save = []
            for item in data:
                flag = task.check_empty_sharepoint(item["folder"])
                if flag is None:
                    continue
                if not flag:
                    num_day = count_day(item["date"])
                    if num_day >= delete_tool_web_day:
                        list_folder.append(item["folder"].split("/")[-1])
                    else:
                        data_save.append(item)
                        if num_day in remind_mail_day:
                            send_reminder_mail(item)

            write_json(check_empty_sharepoint_json, data_save)
            write_file_txt(delete_tool_web, list_folder)
            logger.infor(f"Function check_empty_sharepoint_loop is Successfully!")

            if (time.perf_counter() - start) < time_sleep_check_empty_sharepoint - 1:
                time.sleep(time_sleep_check_empty_sharepoint - (time.perf_counter() - start))

        except Exception as e:
            logger.error(f"Error in check_empty_sharepoint_loop(): {e}")
            logger.error(traceback.format_exc())


def send_email_loop():
    # Initialize COM runtime
    pythoncom.CoInitialize()
    while True:
        if not is_within_runtime_range(end=end_time_n_slot):
            break
        extract_and_sending_email(sending_email_txt)
        time.sleep(email_scan_time)


def count_total_storage_sharepoint_loop():
    pythoncom.CoInitialize()
    show_log_no_data = True
    while True:
        try:
            if not is_within_runtime_range(end=end_time_n_slot):
                break
            start = time.perf_counter()
            data = read_json(count_storage_sharepoint_json)
            if not len(data):
                if show_log_no_data:
                    logger.infor(f"File {count_storage_sharepoint_json} no data!")
                show_log_no_data = False
                time.sleep(30)
                continue
            show_log_no_data = True
            logger.infor(f"File '{count_storage_sharepoint_json}' have data!")

            list_link_45gb = []
            list_link_40gb = []
            task = TaskSharepoint("count_total_storage_sharepoint_loop")

            for item in data:
                link_45gb, link_40gb = task.count_total_storage_sharepoint(item)
                if link_45gb is not None:
                    list_link_45gb.extend(link_45gb)
                if link_40gb is not None:
                    list_link_40gb.extend(link_40gb)

            write_file_txt(reach_45GB_sharepoint_txt, list_link_45gb)
            if len(list_link_40gb):
                write_file_txt(reach_40GB_sharepoint_txt, list_link_40gb)
                for item in list_link_40gb:
                    json_data = {
                        "current_sharepoint": item,
                        "email": mail_admin,
                        "link_request_sharepoint": link_request_sharepoint,
                        "link_sharepoint_management": link_sharepoint_management
                    }
                    send_request_sharepoint_mail(json_data)
                    logger.infor(f"send_emails of count_total_storage_sharepoint_loop to '{mail_admin}' success!")

            logger.infor(f"Function count_total_storage_sharepoint_loop is Successfully!")

            if (time.perf_counter() - start) < time_sleep_count_storage_sharepoint - 1:
                time.sleep(time_sleep_count_storage_sharepoint - (time.perf_counter() - start))

        except Exception as e:
            logger.error(f"Error in count_total_storage_sharepoint_loop(): {e}")
            logger.error(traceback.format_exc())


def gen_folder_in_sharepoint_loop():
    show_log_no_data_gen = True
    show_log_no_data_per = True
    while True:
        try:
            if not is_within_runtime_range(end=end_time_n_slot):
                break
            # start = time.perf_counter()
            data = read_json(new_sharepoint_json)
            if not len(data):
                if show_log_no_data_gen:
                    logger.infor(f"File {new_sharepoint_json} no data!")
                show_log_no_data_gen = False
                time.sleep(5)
            else:
                show_log_no_data_gen = True
                logger.infor(f"File {new_sharepoint_json} have data!")

                task = TaskSharepoint("gen_folder_in_sharepoint_loop")
                for item in data:
                    task.gen_100_folder(item)

                logger.infor(f"Function gen_folder_in_sharepoint_loop is Successfully!")

                # if (time.perf_counter() - start) < time_sleep_new_sharepoint - 1:
                #     time.sleep(time_sleep_new_sharepoint - (time.perf_counter() - start))

                time.sleep(time_sleep_new_sharepoint_and_permission)

            data_ = read_json(permission_group_json)
            if not len(data_):
                if show_log_no_data_per:
                    logger.infor(f"File {permission_group_json} no data!")
                show_log_no_data_per = False
                time.sleep(15)
                continue
            show_log_no_data_per = True
            logger.infor(f"File {permission_group_json} have data!")
            with sync_playwright() as playwright:
                task = AutoSharepoint("permission_group", playwright)
                logger.infor(f"Permission group is stating...!")
                task.stop_inheriting(data_)
                for item in data_:
                    task.group_sharepoint(item)
                    time.sleep(5)
                task.close()

            time.sleep(30)

        except Exception as e:
            logger.error(f"Error in gen_folder_in_sharepoint_loop(): {e}")
            logger.error(traceback.format_exc())


def check_existed_sharepoint_loop():
    show_log_no_data = True
    while True:
        try:
            if not is_within_runtime_range(end=end_time_n_slot):
                break
            start = time.perf_counter()
            data = read_json(check_existed_sharepoint_json)
            if not len(data):
                if show_log_no_data:
                    logger.infor(f"File '{check_existed_sharepoint_json}' no data!")
                show_log_no_data = False
                time.sleep(10)
                continue
            show_log_no_data = True
            logger.infor(f"File '{check_existed_sharepoint_json}' have data!")
            task = TaskSharepoint("check_existed_sharepoint_loop")

            for item in data:
                path = task.check_folder_shortcut_sharepoint(item)
                if path is not None:
                    task.create_folder(path)

            logger.infor(f"Function check_existed_sharepoint_loop is Successfully!")
            time.sleep(10)
            if (time.perf_counter() - start) < time_sleep_check_existed_sharepoint - 1:
                time.sleep(time_sleep_check_existed_sharepoint - (time.perf_counter() - start))

        except Exception as e:
            logger.error(f"Error in check_existed_sharepoint_loop(): {e}")
            logger.error(traceback.format_exc())


def delete_tool_sharepoint_loop():
    show_log_no_data = True
    while True:
        try:
            if not is_within_runtime_range(end=end_time_n_slot):
                break
            start = time.perf_counter()
            data = read_json(delete_tool_sharepoint_json)
            if not len(data):
                if show_log_no_data:
                    logger.infor(f"{delete_tool_sharepoint_json} no data!")
                show_log_no_data = False
                time.sleep(30)
                continue
            show_log_no_data = True
            logger.infor(f"file '{delete_tool_sharepoint_json}' have data!")
            task = TaskSharepoint("delete_tool_sharepoint_loop")
            for item in data:
                task.delete_tool(item)

            logger.infor(f"Function delete_tool_sharepoint_loop is Successfully!")

            if (time.perf_counter() - start) < time_sleep_delete_tool_sharepoint - 1:
                time.sleep(time_sleep_delete_tool_sharepoint - (time.perf_counter() - start))

        except Exception as e:
            logger.error(f"Error in delete_tool_sharepoint_loop(): {e}")
            logger.error(traceback.format_exc())


def move_data_sharepoint_loop():
    show_log_no_data = True
    while True:
        try:
            if not is_within_runtime_range(end=end_time_n_slot):
                break
            start = time.perf_counter()
            data = read_json(move_folder_json)
            if not len(data):
                if show_log_no_data:
                    logger.infor(f"File '{move_folder_json}' no data!")
                show_log_no_data = False
                time.sleep(15)
                continue
            show_log_no_data = True
            logger.infor(f"File '{move_folder_json}' have data!")
            task = TaskSharepoint("move_data_sharepoint_loop")
            for item in data:
                task.move_data(item)

            logger.infor(f"Function move_data_sharepoint_loop is Successfully!")

            if (time.perf_counter() - start) < time_sleep_move_folder - 1:
                time.sleep(time_sleep_move_folder - (time.perf_counter() - start))

        except Exception as e:
            logger.error(f"Error in move_data_sharepoint_loop(): {e}")
            logger.error(traceback.format_exc())


def update_current_user_loop():
    pythoncom.CoInitialize()
    show_log_no_data = True
    while True:
        try:
            if not is_within_runtime_range(end=end_time_n_slot):
                break
            start = time.perf_counter()
            data = read_file_txt(update_current_user_txt)
            web_import_account = "web-import-account"
            if not len(data) or data[0] == web_import_account:
                if show_log_no_data:
                    logger.infor(f"File '{update_current_user_txt}' doesn't want to process!")
                show_log_no_data = False
                time.sleep(15)
                continue
            elif data[0] == "python-check-account":
                write_file_txt(update_current_user_txt, [web_import_account])
                show_log_no_data = True
                task = UserOutlook("Update_current_user")
                task.check_user()
                write_file_txt(update_current_user_txt, [web_import_account])

                logger.infor(f"Update_current_user_loop is Successfully!")
                print("Update_current_user_loop: ", time.perf_counter() - start)
                # time.sleep(100)

                if (time.perf_counter() - start) < time_sleep_move_folder - 1:
                    time.sleep(time_sleep_move_folder - (time.perf_counter() - start))
            else:
                time.sleep(10)

        except Exception as e:
            logger.error(f"Error in Update_current_user_loop(): {e}")
            logger.error(traceback.format_exc())


if __name__ == '__main__':
    state, message = set_startup(PROJECT_NAME, PROJECT_REMINDER_EXE_PATH)

    mount_network_drive()
    notify_setup_success()
    thread_1 = None
    thread_2 = None
    thread_3 = None
    thread_4 = None
    thread_5 = None
    thread_6 = None
    thread_7 = None
    thread_8 = None
    thread_9 = None
    while True:
        if is_within_runtime_range(start="4:45:00", end="4:55:00"):
            thread_1 = None
            thread_2 = None
            thread_3 = None
            thread_4 = None
            thread_5 = None
            thread_6 = None
            thread_7 = None
            thread_8 = None
            thread_9 = None

        if is_within_runtime_range(start=start_time_one_slot, end=set_end_time(start_time_one_slot)):
            logger.infor(f"One slot start action!")
            if thread_1 is None:
                thread_1 = threading.Thread(target=delete_unuse_folder_loop)
                thread_1.start()
                time.sleep(1)

            if thread_2 is None:
                thread_2 = threading.Thread(target=check_empty_sharepoint_loop)
                thread_2.start()
                time.sleep(1)
            pass
        else:
            if thread_1 is not None:
                thread_1 = None
            if thread_2 is not None:
                thread_2 = None
            pass

        if is_within_runtime_range(end=end_time_n_slot):
            # if thread_3 is None:
            #     thread_3 = threading.Thread(target=send_email_loop)
            #     thread_3.start()
            #     time.sleep(1)
            #
            # if thread_4 is None:
            #     thread_4 = threading.Thread(target=count_total_storage_sharepoint_loop)
            #     thread_4.start()
            #     time.sleep(1)

            if thread_5 is None:
                thread_5 = threading.Thread(target=gen_folder_in_sharepoint_loop)
                thread_5.start()
                time.sleep(1)

            # if thread_6 is None:
            #     thread_6 = threading.Thread(target=check_existed_sharepoint_loop)
            #     thread_6.start()
            #     time.sleep(1)
            #
            # if thread_7 is None:
            #     thread_7 = threading.Thread(target=delete_tool_sharepoint_loop)
            #     thread_7.start()
            #     time.sleep(1)
            #
            # if thread_8 is None:
            #     thread_8 = threading.Thread(target=move_data_sharepoint_loop)
            #     thread_8.start()
            #     time.sleep(1)
            #
            # if thread_9 is None:
            #     thread_9 = threading.Thread(target=update_current_user_loop)
            #     thread_9.start()
            #     time.sleep(1)

        else:
            logger.infor(f"N_slot not in time action!")

            if thread_3 is not None:
                thread_3 = None

            if thread_4 is not None:
                thread_4 = None

            if thread_5 is not None:
                thread_5 = None

            if thread_6 is not None:
                thread_6 = None

            if thread_7 is not None:
                thread_7 = None

            if thread_8 is not None:
                thread_8 = None

            if thread_9 is not None:
                thread_9 = None
            time.sleep(waiting_5_minute)
