import win32com.client
import pandas as pd
import utils
from configs import *


class UserOutlook:
    def __init__(self, function_loop):
        self.function_loop = function_loop

        self.olApp = win32com.client.Dispatch("Outlook.Application")

        self.olNamespace = self.olApp.GetNamespace("MAPI")

    def check_user(self):
        logger.infor(f"{self.function_loop} is processing!")

        data_user_input = utils.read_csv(input_check_user_outlook_csv)
        knt_list = data_user_input["KNT"]
        data_config = utils.read_csv(config_check_user_outlook_csv)

        dept_dict = self.check_data_config(data_config)
        user_knt_list, department_list, dept_list, mail_list = self.check_data_input(knt_list, dept_dict)
        self.save_data(user_knt_list, department_list, dept_list, mail_list)

    def check_data_input(self, knt_list, dept_dict):
        department_list = []
        user_knt_list = []
        dept_list = []
        mail_list = []

        for i in range(len(knt_list)):
            recipient = self.olNamespace.CreateRecipient(knt_list[i])
            address_entry = recipient.AddressEntry
            try:
                user = address_entry.GetExchangeUser()
                mail = user.PrimarySmtpAddress
                dept = user.Department
                user_knt_list.append(knt_list[i])
                dept_list.append(dept)
                mail_list.append(mail)
                department_list.append(dept_dict[dept] if dept in dept_dict.keys() else "")
            except Exception as e:
                print(e)
                continue
        return user_knt_list, department_list, dept_list, mail_list

    def check_data_config(self, data_config):
        dept_dict = {}
        for i in range(len(data_config["username"])):
            recipient = self.olNamespace.CreateRecipient(data_config["username"][i])
            address_entry = recipient.AddressEntry
            try:
                user = address_entry.GetExchangeUser()
                dept = user.Department
                dept_dict[dept] = data_config.department[i]
            except Exception as e:
                print(e)
                continue
        return dept_dict

    @staticmethod
    def save_data(user_knt_list, department_list, dept_list, mail_list):
        data_output = pd.DataFrame()
        data_output["KNT"] = user_knt_list
        data_output["Department"] = department_list
        data_output["DEPT"] = dept_list
        data_output["EMAIL"] = mail_list
        data_output = data_output.astype(str)
        data_output.to_csv(output_check_user_outlook_csv)


if __name__ == "__main__":
    task = UserOutlook("Check user outlook")
    path_config = r"C:\Users\KNT21818\Downloads\CHECK_KNT 1\CHECK_KNT\config.csv"
    path_input = r"C:\Users\KNT21818\Downloads\CHECK_KNT 1\CHECK_KNT\File_output.csv"
    task.check_user()
