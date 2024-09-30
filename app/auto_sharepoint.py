import os.path
import time

from configs import *
from utils import read_csv, write_file_txt


class AutoSharepoint:
    def __init__(self, function_loop, playwright):
        self.function_loop = function_loop
        self.flag_login = True
        self.browser = playwright.chromium.launch(channel='msedge', headless=False)
        self.context = self.browser.new_context()

        self.item = None
        self.page = None

    def go_to_web(self, document=False):
        time.sleep(5)
        if document:
            self.page.goto("https://nissangroup.sharepoint.com/teams/" + self.item[
                "link"] + "/Shared%20Documents/Forms/AllItems.aspx")

        else:
            self.page.goto("https://nissangroup.sharepoint.com/teams/" + self.item["link"] + "/_layouts/15/user.aspx")
        self.page.wait_for_timeout(40000)

    def login_knt(self, knt):
        self.page.fill("[name='loginfmt']", knt)
        self.page.press("[name='loginfmt']", "Enter")
        self.page.wait_for_timeout(30000)
        time.sleep(3)

    def step_click_create_group(self):
        self.page.click("a[id='Ribbon.Permission.Add.NewGroup-Large']")
        self.page.wait_for_timeout(5000)

    def step_enter_name_group(self):
        self.page.fill("input[title='Name of the group']", self.item["group_name"])
        self.page.wait_for_timeout(5000)

    def step_set_permission(self):
        self.page.click("label:has-text('Read - Can view pages and list items and download documents.')")
        self.page.wait_for_timeout(5000)

    def step_click_create(self):
        self.page.click("input[value='Create']")
        self.page.wait_for_timeout(5000)

    def step_click_new(self):
        try:
            self.page.click("a:has-text('New')")
            self.page.wait_for_timeout(5000)
        except Exception as e:
            print(e)
            self.page.goto("https://nissangroup.sharepoint.com/teams/" + self.item[
                "link"] + "/_layouts/15/groups.aspx")
            self.page.wait_for_timeout(10000)
            self.page.click(f"""a:has-text('{self.item["group_name"]}')""")
            self.page.wait_for_timeout(5000)
            self.step_click_new()

    def step_enter_mail(self, list_mail):
        frame = self.page.frames[3]
        frame.click("a:has-text('Show options')")
        self.page.wait_for_timeout(5000)
        frame.click("label:has-text('Send an email invitation')")
        self.page.wait_for_timeout(5000)
        frame.click("div[title='Enter names or email addresses.']")
        self.page.wait_for_timeout(1000)
        self.page.evaluate(f"""navigator.clipboard.writeText('{list_mail}')""")

        self.page.keyboard.press("Control+V")
        time.sleep(15)
        self.page.wait_for_timeout(5000)
        frame.click("input[value='Share']")
        self.page.wait_for_timeout(10000)

    def step_click_mouse_right_folder(self, folder):
        self.page.click(f"button[title='{folder}']", button="right")
        self.page.wait_for_timeout(5000)

    def step_click_detail(self):
        self.page.click("span:has-text('Details')")
        self.page.wait_for_timeout(6000)

    def step_click_manage_access(self):
        self.page.click("span:has-text('Manage access')")
        self.page.wait_for_timeout(5000)

    def step_click_share(self):
        self.page.click("button[name='Share']")
        self.page.wait_for_timeout(6000)

    def step_add_name_group(self):
        share_frame = self.page.wait_for_selector("iframe#shareFrame", state="visible")
        frame = share_frame.content_frame()
        self.step_set_view(frame)
        frame.click('input[aria-label="Add a name, group, or email"]')
        self.page.wait_for_timeout(5000)
        self.page.evaluate(f"""navigator.clipboard.writeText('{self.item["group_name"]}')""")
        self.page.wait_for_timeout(5000)

        self.page.keyboard.press("Control+V")
        self.page.wait_for_timeout(5000)
        self.page.keyboard.press("Enter")
        self.page.wait_for_timeout(5000)
        frame.click('button:has-text("Send")')
        self.page.wait_for_timeout(5000)

    def step_set_view(self, frame):
        frame.click("i[data-icon-name='Edit']")
        self.page.wait_for_timeout(5000)
        frame.click("span:has-text('Can view')")
        self.page.wait_for_timeout(5000)

    def group_sharepoint(self, item):
        try:
            self.item = item
            if self.item["group_name"] is '':
                return
            self.process_file_email()
            logger.infor(f"create group {self.item['link']} is processing !")

            self.go_to_web()
            self.step_click_create_group()
            self.step_enter_name_group()
            self.step_set_permission()
            self.step_click_create()
            for idx in range(0, len(self.item["email"]), 100):
                try:
                    self.step_click_new()
                    list_email = ";".join(self.item["email"][idx:idx + 100])
                    self.step_enter_mail(list_email)
                except Exception as e:
                    print(e)
                    path_save_error = os.path.join(input_folder, f"/error_mail/error_add_email_{time.perf_counter()}")
                    write_file_txt(path_save_error, self.item["email"][idx:idx + 100])
                    logger.infor(
                        f"Add number error start folder{path_save_error} !")
                    continue
            logger.infor(f"create group {self.item['link']} is Successfully !")
        except Exception as e:
            print(e)
            pass

        # for idx, path_folder in enumerate(self.item["folder"]):
        #     try:
        #         logger.infor(f"add group sharepoint is processing sharepoint: {self.item['link']}/{path_folder}!")
        #         self.go_to_web(document=True)
        #         self.step_close_info()
        #         list_folder = path_folder.split('/')
        #         self.step_go_to_folder(list_folder[:len(list_folder) - 1])
        #         self.step_click_mouse_right_folder(list_folder[-1])
        #         self.step_click_share()
        #         self.step_add_name_group()
        #         logger.infor(f"add group is successful!--- {self.item['link']}/{path_folder}!")
        #
        #     except Exception as e:
        #         print(e)
        #         continue
        logger.infor(f"Permission group is Successfully sharepoint: {self.item['link']}!")
        time.sleep(5)

    def step_click_more(self):
        share_frame = self.page.wait_for_selector("iframe#shareFrame", state="visible")
        frame = share_frame.content_frame()
        self.page.wait_for_timeout(5000)
        frame.click("i[data-icon-name='More']")
        time.sleep(1)
        frame.click("button:has-text('Advanced settings')")
        self.page.wait_for_timeout(5000)

    def step_click_stop_inheriting_permissions(self):

        time.sleep(5)
        new_page = self.context.pages[-1]
        new_page.bring_to_front()
        try:
            self.page.wait_for_timeout(1000)
            new_page.on("dialog", lambda dialog: dialog.accept())
            new_page.click("a[id='Ribbon.Permission.Manage.StopInherit-Large']")
            new_page.wait_for_timeout(5000)
        except Exception as e:
            print(e)
            pass
        new_page.close()
        self.context.pages[0].bring_to_front()

    def step_go_to_folder(self, list_folder):
        for folder in list_folder:
            if folder != '':
                self.page.click(f"""button[title='{folder}']""")
                self.page.wait_for_timeout(5000)

    def step_close_info(self):
        self.page.click("i[data-icon-name='Info']")
        self.page.wait_for_timeout(5000)

    def stop_inheriting(self, data):
        for it in data:
            self.item = it
            self.page = self.context.new_page()

            self.go_to_web(document=True)
            if self.flag_login:
                self.login_knt("KNT21818@local.nmcorp.nissan.biz")
            else:
                page = self.context.pages[0]
                page.close()
            self.flag_login = False
            for idx, path_folder in enumerate(self.item["folder"]):
                logger.infor(f"Stop inheriting is processing sharepoint: {self.item['link']}/{path_folder}!")
                if idx != 0:
                    self.go_to_web(document=True)
                    self.step_close_info()
                try:
                    list_folder = path_folder.split('/')
                    self.step_go_to_folder(list_folder[:len(list_folder) - 1])
                    self.step_click_mouse_right_folder(list_folder[-1])
                    self.step_click_manage_access()
                    self.step_click_more()
                    self.step_click_stop_inheriting_permissions()
                    time.sleep(5)
                except Exception as e:
                    print(e)
                    continue

    def process_file_email(self):
        path_excel = os.path.join(input_folder, export_user_email_csv)
        df = read_csv(path_excel)
        list_mail = []
        list_dept_email_natv = email_NATV.split(",")
        list_dept_email_nat = email_NAT.split(",")
        if "and" in self.item["group_name"]:
            for col in list_dept_email_natv:
                list_mail.extend(list(df[col]))
            for idx, col in enumerate(list_dept_email_nat):
                list_mail.extend(list(df[col]))
        elif "NATV" in self.item["group_name"]:
            for idx, col in enumerate(list_dept_email_natv):
                list_mail.extend(list(df[col]))
        elif "NAT" in self.item["group_name"]:
            for idx, col in enumerate(list_dept_email_nat):
                list_mail.extend(list(df[col]))
        elif self.item["group_name"] is not '':
            list_mail = list(df[self.item["group_name"].split('-')[0]])
        list_mail = [x for x in list(set(list_mail)) if str(x) != 'nan']
        self.item["email"] = list_mail

    def close(self):
        self.browser.close()

# if __name__ == "__main__":
#     from playwright.sync_api import sync_playwright
#     items = [
#         {
#             "link": "JAO_NTV_000402_NATV-ADM-Automation-Tools-01",
#             "folder": [
#                 "AG0/NAT_NATV_Public"
#             ],
#             "group_name": "NATV and NAT",
#
#
#         }
#     ]
#     with sync_playwright() as p:
#         task = AutoSharepoint("pass", p)
#         task.stop_inheriting(items[0])
#         task.group_sharepoint()
