# ↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓
# Created by Tien Dat(KNT19856) on 2023/10/16 2:56 pm
# Authors: 
# Group: DX
# Project: 
# Script version: 
# Last modify: 
# Description: None
# ↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑

# Khởi động cùng máy tính

import winreg as wr


def set_startup(program_name, program_path):
    try:
        key = wr.HKEY_CURRENT_USER
        sub_key = "Software\\Microsoft\\Windows\\CurrentVersion\\Run"
        program_name = program_name  # Tên của chương trình khởi động cùng máy tính
        program_path = program_path

        with wr.OpenKey(key, sub_key, 0, wr.KEY_WRITE) as reg_key:
            wr.SetValueEx(reg_key, program_name, 0, wr.REG_SZ, f'"{program_path}"')

        return True, f"setStartup({program_name}, {program_path}) Successfully!"
    except Exception as e:
        return False, f"setStartup({program_name}, {program_path}) Error! {e}"


def remove_startup(program_name):
    try:
        key = wr.HKEY_CURRENT_USER
        sub_key = "Software\\Microsoft\\Windows\\CurrentVersion\\Run"
        program_name = program_name  # Tên của chương trình khởi động cùng máy tính

        with wr.OpenKey(key, sub_key, 0, wr.KEY_WRITE) as reg_key:
            wr.DeleteValue(reg_key, program_name)

        return True, f"removeStartup({program_name}) Successfully!"
    except Exception as e:
        return False, f"removeStartup({program_name}) Error! {e}"

# if __name__ == '__main__':
#     setStartup(program_name="Project_Reminder",
#     program_path=r"C:\Users\KNT19856\Desktop\python\Project_Reminder\src\GUI\pyuic5.exe")
# removeStartup(program_name="Project_Reminder")
