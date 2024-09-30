import logging
from datetime import datetime
from pathlib import Path
import os
import shutil


class OperationLog:
    def __init__(self, save_folder=None, log_life_circle=180, log_name="Operation Log"):
        self.log_life_circle = log_life_circle
        self.save_folder = save_folder
        self.log_name = log_name
        self.logger = None
        if not self.save_folder:
            self.save_folder = os.path.join(os.path.dirname(__file__), "Logs")

        # Check and create folder if not
        Path(self.save_folder).mkdir(parents=True, exist_ok=True)
        # Config Log
        self.config_log(flag=False)
        try:
            self.delete_expired_folder()
        except Exception as err_name:
            self.warning(err_name)
            pass

    def config_log(self, flag=True):
        now_day = datetime.now().strftime('%Y-%m-%d')
        if now_day not in self.log_name or not flag:
            self.log_name = f"Log [{now_day}]"
            self.logger = logging.getLogger(self.log_name)
            formatter = logging.Formatter(f'%(message)s')
            file_handler = logging.FileHandler(f"{self.save_folder}/{self.log_name}.log", mode='a', encoding="utf-8")
            file_handler.setFormatter(formatter)
            stream_handler = logging.StreamHandler()
            stream_handler.setFormatter(formatter)
            self.logger.setLevel(logging.DEBUG)
            self.logger.addHandler(file_handler)
            self.logger.addHandler(stream_handler)

    def debug(self, content):
        self.config_log()
        now = datetime.now()
        current_date_time = now.strftime("%Hh-%Mm-%S")
        self.logger.debug(f"{current_date_time}-[DEBUG]: {content}\n")

    def infor(self, content):
        self.config_log()
        now = datetime.now()
        current_date_time = now.strftime("%Hh-%Mm-%S")
        self.logger.info(f"{current_date_time}-[INFO]: {content}\n")

    def warning(self, content):
        self.config_log()
        now = datetime.now()
        current_date_time = now.strftime("%Hh-%Mm-%S")
        self.logger.warning(f"{current_date_time}-[WARNING]: {content}\n")

    def error(self, content):
        self.config_log()
        now = datetime.now()
        current_date_time = now.strftime("%Hh-%Mm-%S")
        self.logger.error(f"{current_date_time}-[ERROR]: {content}\n")

    # def screenshot_error_log(self, screenshot_name=""):
    #     now = datetime.now()
    #     current_date_time = now.strftime("%Hh%Mm%S")
    #     save_path = self.save_folder + f"/{current_date_time} {screenshot_name}.jpg"
    #     pyautogui.screenshot(save_path)
    #     self.error(f"Saved screenshot in {save_path}")

    def delete_expired_folder(self):
        self.config_log()
        # Input
        folder = os.path.dirname(self.save_folder)
        day = self.log_life_circle

        # Check and delete all folder older than day
        folder_paths_list = [f.path for f in os.scandir(folder) if f.is_dir()]
        for folder in folder_paths_list:
            folder_last_modified_time = max(os.stat(root).st_mtime for root, _, _ in os.walk(folder))
            # convert timestamp to datetime
            modification_date = datetime.fromtimestamp(folder_last_modified_time)
            # find the number of days when the file was modified
            number_of_days = (datetime.now() - modification_date).days
            if number_of_days >= day:
                try:
                    shutil.rmtree(folder)
                    self.infor(f"Logs folder: '{folder}' has been deleted")
                except Exception as err:
                    # self.warning(err)
                    self.warning(f"Directory '% s' can not be removed. {err}" % folder)

# if __name__ == "__main__":
#     logger1 = OperationLog(log_name="Test1")
#     logger1.infor("test 1 info")
#
#     logger2 = OperationLog(log_name="Test2")
#     logger2.infor("test 2 info")
