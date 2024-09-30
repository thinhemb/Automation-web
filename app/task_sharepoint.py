
import shutil
import itertools
from pathlib import Path
import time
import copy

from configs import *


class TaskSharepoint:
    def __init__(self, function_loop):
        self.function_loop = function_loop

    @staticmethod
    def get_folder_id_max(folder_sharepoint, folder):
        path_folder_sharepoint = os.path.join(folder_sharepoint, folder)
        list_name_folder = os.listdir(path_folder_sharepoint)
        groups = itertools.groupby(list_name_folder, key=lambda item: item[item.index("_"):item.rfind("_")])
        # Chuyển nhóm thành list các list
        res_list = [list(group) for key, group in groups]
        res_id = []
        for group in res_list:
            res_id.append(os.path.join(folder, group[-1]))
        return ",".join(res_id)

    @staticmethod
    def get_capacity_sharepoint(folder_sharepoint):
        total_size = 0
        for root, _, files in os.walk(folder_sharepoint):
            # Duyệt qua tất cả các file trong folder
            for file in files:
                # Lấy đường dẫn file
                file_path = os.path.join(root, file)

                file_size = os.path.getsize(u"\\\\?\\" + file_path)
                # Cộng dồn kích thước file vào tổng dung lượng
                total_size += file_size
        return total_size / 1024 ** 3

    def count_total_storage_sharepoint(self, item):
        folder_sharepoint = self.check_folder_shortcut_sharepoint(item["folder-sharepoint"])
        if folder_sharepoint is None or not self.check_exist_folder(folder_sharepoint):
            return None, None
        logger.infor(f"{self.function_loop} is processing sharepoint: {item['folder-sharepoint']}!")
        capacity = self.get_capacity_sharepoint(folder_sharepoint)
        list_link_45gb = []
        list_link_40gb = []
        # Định dạng thời gian theo form "%Y/%m/%d %H:%M:%S"
        time_str = datetime.now().strftime("%Y/%m/%d %H:%M:%S")

        if capacity > reach_45GB_sharepoint_capacity:
            for path in item["folder"]:
                folder_id_max = self.get_folder_id_max(folder_sharepoint, path)
                # print(item["link"] + " " + folder_id_max + " " + time_str)

                link = item["folder-sharepoint"] + " " + folder_id_max + " " + time_str
                list_link_45gb.append(link)
        elif reach_40GB_sharepoint_capacity < capacity < reach_45GB_sharepoint_capacity:
            link = item["folder-sharepoint"] + " " + time_str
            if link not in list_link_40gb:
                list_link_40gb.append(link)

        return list_link_45gb, list_link_40gb

    def gen_100_folder(self, item):
        path_folder_new = self.check_folder_shortcut_sharepoint(item["folder-sharepoint"])
        if path_folder_new is None or not self.check_exist_folder(path_folder_new):
            return
        logger.infor(f"{self.function_loop} is processing sharepoint: {item['folder-sharepoint']}!")
        for path in item["folder"]:

            list_name_folder = path.split("/")

            name_folder_last = list_name_folder[-1]
            position = name_folder_last.rfind("_")
            first_name = name_folder_last[:position + 1]
            start_index = int(name_folder_last[position + 1:])
            for idx, flag in enumerate(flag_folder):
                if idx == 1:
                    flag = item["company"] + "_Public"
                data = self.add_before_fy(list_name_folder, flag)
                path_folder_contain = os.path.join(path_folder_new,
                                                   "/".join(data[:len(data) - 1]))
                for index in range(start_index, start_index + 100):
                    path_folder = os.path.join(path_folder_contain, first_name + format(index, '03'))
                    self.create_folder(path_folder)

    @staticmethod
    def add_before_fy(data, new_flag):
        new_data = copy.deepcopy(data)
        for i, element in enumerate(new_data):
            for flag in flag_folder:
                if flag in new_data:
                    return new_data
            if "FY" in element[:2]:
                new_data.insert(i, new_flag)
                break
        return new_data

    def delete_tool(self, path):
        path_folder_new = self.check_folder_shortcut_sharepoint(path)
        if path_folder_new is None or not self.check_exist_folder(path_folder_new):
            return
        logger.infor(f"{self.function_loop} is processing sharepoint: {path}!")
        shutil.rmtree(path_folder_new)

    def delete_folder(self, item):
        path_folder_new = self.check_folder_shortcut_sharepoint(item["folder-sharepoint"])
        if path_folder_new is None or not self.check_exist_folder(path_folder_new):
            return
        logger.infor(f"{self.function_loop} is processing sharepoint: {item['folder-sharepoint']}!")
        for path in item["folder"]:

            list_name_folder = path.split("/")
            name_folder_last = list_name_folder[-1]
            position = name_folder_last.rfind("_")
            first_name = name_folder_last[:position + 1]
            start_index = int(name_folder_last[position + 1:])
            path_folder_contain = os.path.join(path_folder_new, "/".join(list_name_folder[:len(list_name_folder) - 1]))
            for index in range(start_index, start_index + 100):
                path_folder = os.path.join(path_folder_contain, first_name + format(index, '03'))
                if self.check_exist_folder(path_folder):
                    if len(os.listdir(path_folder)) > 0:
                        continue
                    os.rmdir(path_folder)

    def check_empty_sharepoint(self, path):
        path_folder_new = self.check_folder_shortcut_sharepoint(path)
        if path_folder_new is None:
            return None
        logger.infor(f"{self.function_loop} is processing sharepoint: {path}!")

        flag = self.check_exist_folder(path_folder_new)
        if not flag:
            return None
        list_file = os.listdir(path_folder_new)
        if len(list_file) == 0:
            return False
        return True

    @staticmethod
    def check_exist_folder(path):
        return os.path.exists(path)

    def check_folder_shortcut_sharepoint(self, path):

        name_pc = os.getlogin()
        path = path.replace("\\", "/")
        path_split = path.split("/")
        path_onedrive = r"C:\Users\{}\OneDrive - Nissan Motor Corporation".format(name_pc)
        path_temp = os.path.join(path_onedrive, "Shared Documents - " + path_split[0])
        if self.check_exist_folder(path_temp):
            return os.path.abspath(os.path.join(path_temp, "/".join(path_split[1:])))

        path_temp = os.path.join(path_onedrive, "Documents - " + path_split[0])
        if self.check_exist_folder(path_temp):
            return os.path.abspath(os.path.join(path_temp, "/".join(path_split[1:])))

        logger.infor(f"sharepoint shortcut {path_temp} is not exist!")
        return None

    @staticmethod
    def create_folder(folder_path):
        # os.makedirs(folder_path, exist_ok=True)
        Path(folder_path).mkdir(parents=True, exist_ok=True)

    def move_data(self, item):
        path_old = self.check_folder_shortcut_sharepoint(item["old"])
        path_new = self.check_folder_shortcut_sharepoint(item["new"])
        if path_new is None:
            return
        self.create_folder(path_new)
        if path_old is None or not self.check_exist_folder(path_old) or path_old == path_new:
            return
        else:
            logger.infor(f"{self.function_loop} is processing sharepoint: {item['old']}!")
            # if not os.path.exists(path_new):

            shutil.copytree(path_old, path_new, dirs_exist_ok=True)
            time.sleep(60)
            shutil.rmtree(path_old)

# if __name__ == '__main__':
#     link = r''
#     task = TaskSharepoint()
#
#     start = time.perf_counter()
#     # print(cap.get_capacity_sharepoint(link))
#     task.go_to_web(link)
#     # task.gen_100_folder('test_sharepoint', '2023_DR0_0')
#     # print(task.delete_tool(['test_sharepoint', '2023_DR0_6']))
#     #
#     print(task.delete_folder(['test_sharepoint', '2023_DR0_60']))
#     print(time.perf_counter() - start)
#     task.close()
