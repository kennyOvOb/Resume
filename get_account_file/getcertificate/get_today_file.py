import pandas as pd

import shutil
import logging
from datetime import timedelta, datetime
from path import DESKTOP, GA_BASE_ACCOUNT_PUBLIC_DIR, OVER_BOARD_PATH, \
    CERTIFICATE_PATH
from ..account_data.account_day import get_account_day
from ..account_data.distribution import Distribution
from ..account_data.database import Database
from ..common_functions.functions import get_update_file, get_last_day


class TodayFile:
    def __init__(self, str_date_ymd=get_account_day()[5]):
        self.str_date_ymd = str_date_ymd
        self.account_day = datetime.strptime(str_date_ymd, '%y%m%d')
        self.distribution = Distribution(self.account_day)
        self.database = Database()

    def add_log(self, folder_name, file_name, log):
        site_note_dict = self.distribution.get_distribute_note()
        log_df: pd.DataFrame = pd.DataFrame(log)
        if folder_name == "問題整合":
            log_df.columns = ["客戶", "結果"]
        else:
            log_df.columns = ["客戶", "檔案日期", "結果"]
        log_df["備註"] = log_df["客戶"].str.upper().map(site_note_dict)
        log_df.to_excel(self.get_storage_path(folder_name=folder_name) / file_name, index=False)

    @staticmethod
    def get_storage_path(folder_name, create_path=DESKTOP):  # 抓檔案放置地方，預設是放桌面
        storage_path = create_path / folder_name
        return storage_path

    @staticmethod
    def get_account_path(account_day_str):  # 取得ga上彙總放置路徑
        account_path = GA_BASE_ACCOUNT_PUBLIC_DIR / ("帳務日" + account_day_str) / "帳務"
        return account_path

    def get_last_account_day_in_ga_base_account(self):
        ga_base_account_folders = filter(
            lambda file: ("帳務日" + self.str_date_ymd[2:4] in file.name) and (file.name != "帳務日(空白)"),
            GA_BASE_ACCOUNT_PUBLIC_DIR.glob("*"))
        day_list = [datetime(int(self.str_date_ymd[:2]), int(folder.name.replace("帳務日", "")[:2]),
                             int(folder.name.replace("帳務日", "")[-2:])) for folder in ga_base_account_folders]
        if day_list:
            min_date_by_day = min(day_list).day
        else:
            min_date_by_day = get_last_day(self.str_date_ymd) + 1
        return min_date_by_day

    @staticmethod
    def creat_folder(folder_name, create_path=DESKTOP):  # 需要創建的資料夾名稱，路徑預設放在桌面
        if not (create_path / folder_name).is_dir():
            (create_path / folder_name).mkdir()

    @staticmethod
    def get_whole_file_name(filename, filename_extension=".xlsx"):  # 檔案名稱包含副檔名，預設用xlsx
        whole_file_name = filename + filename_extension
        return whole_file_name

    # 關帳模板
    @staticmethod
    def get_over_mould_board_path(over_mould_board_name):
        return OVER_BOARD_PATH / (over_mould_board_name + ".xlsx")

    def move_over_mould_board(self, file_needed_list, log):
        storage_path = self.get_storage_path(folder_name="關帳模板")  # 儲存路徑
        name_number_dict = self.distribution.name_number_dict
        if not storage_path.exists():
            self.creat_folder(folder_name="關帳模板")
        for client in file_needed_list:
            if client not in name_number_dict:
                log.append([client, "無此模板", "複製失敗"])
                continue
            over_mould_board_name = name_number_dict[client] + "_" + client
            over_mould_board_path = self.get_over_mould_board_path(over_mould_board_name)
            if over_mould_board_path.exists():
                try:
                    shutil.copy(over_mould_board_path, storage_path / (client + ".xlsx"))
                    log.append([client, "有此模板", "複製成功"])
                except Exception as msg:  # noqa
                    print(msg)
                    log.append([client, "有此模板", "複製失敗"])
            else:
                log.append([client, "無此模板", "複製失敗"])

        return log

    # 關帳彙總

    def get_over_summary_path(self, client):
        date = datetime.strptime(self.str_date_ymd, "%y%m%d")

        return CERTIFICATE_PATH / (str(date.year) + "-" + str(date.month)) / client / "7.报表凭证"

    @staticmethod
    def looking_for_summary(all_summary_filter, filter_str_date_ymd):
        summary_filter = filter(lambda file: filter_str_date_ymd in file.stem, all_summary_filter)  # copy不知道要不要加
        summary_filter_list = list(summary_filter)
        if len(summary_filter_list) == 0:  # 代表沒找到
            return None
        elif len(summary_filter_list) == 1:  # 代表僅有一個
            return summary_filter_list[0]
        else:
            return get_update_file(summary_filter_list)  # 代表多個，要找到正確的

    def move_over_summary(self, file_needed_list, log):
        storage_path = self.get_storage_path(folder_name="關帳彙總")  # 儲存路徑
        if not storage_path.exists():
            self.creat_folder(folder_name="關帳彙總")
        for client in file_needed_list:
            over_summary_path = self.get_over_summary_path(client)
            all_summary_filter = filter(
                lambda file: client.upper() in file.stem.upper() and "汇总" in file.stem and file.stat().st_size > 5000
                             and "~$" not in file.stem and "._" not in file.stem,
                over_summary_path.glob("*"))
            all_summary_list = list(all_summary_filter)
            account_day_by_day = int(self.str_date_ymd[-2:])
            account_day = datetime.strptime(self.str_date_ymd, "%y%m%d").date()
            filter_str_date_ymd = self.str_date_ymd
            while account_day_by_day > 0:
                summary_path = self.looking_for_summary(all_summary_list, filter_str_date_ymd)
                if summary_path:  # 確認是否有檔案
                    # 叫pycharm安靜
                    # noinspection PyBroadException
                    try:
                        shutil.copy(summary_path, storage_path / (client + "-汇总" + summary_path.suffix))
                        log.append([client, account_day, "Success"])
                        break
                    except Exception:  # 不知道例外處理方式，如果有檔案卻抓不出來，記錄下
                        log.append([client, account_day, "Copy Failed"])
                        break
                else:
                    account_day_by_day -= 1
                    account_day -= timedelta(days=1)
                    if account_day_by_day == 0:
                        log.append([client, "本月無此彙總", "Failed"])  # 如果日期歸零，代表沒模板
                        break
                    else:
                        filter_str_date_ymd = account_day.strftime("%y%m%d")
        return log

    # 平日模板
    @staticmethod
    def get_file_needed_list(path):  # 需要抓取檔案的excel表
        file_needed = pd.read_excel(path).iloc[:, :1].dropna()  # 開啟檔案取得需要模板的名稱
        column_name = file_needed.columns[0]
        file_needed_list = file_needed[column_name].tolist()  # 做成list 以便循環抓取
        return file_needed_list

    def get_mould_board_path(self, account_day_str, whole_file_name, account_day_by_day):  # 取得ga上模板放置路徑
        last_account_day_in_ga_base_account = self.get_last_account_day_in_ga_base_account()
        if account_day_by_day >= last_account_day_in_ga_base_account:
            mould_board_path = GA_BASE_ACCOUNT_PUBLIC_DIR / ("帳務日" + account_day_str) / "模板" / whole_file_name  #
        else:
            mould_board_path = GA_BASE_ACCOUNT_PUBLIC_DIR / (
                datetime.strptime(self.str_date_ymd, "%y%m%d").strftime("%Y.%#m")) / (
                                       "帳務日" + account_day_str) / "模板" / whole_file_name  #
        return mould_board_path

    def move_mould_board(self, file_needed_list, log):
        storage_path = self.get_storage_path(folder_name="當日模板")  # 儲存路徑
        if not storage_path.exists():
            self.creat_folder(folder_name="當日模板")
        for client in file_needed_list:
            whole_file_name = self.get_whole_file_name(client)  # 完整檔案包含副檔案名稱
            account_day_by_day = int(self.str_date_ymd[-2:])
            account_day_str = self.str_date_ymd[-4:]
            account_day = datetime.strptime(self.str_date_ymd, "%y%m%d").date()
            while account_day_by_day > 0:
                #  完整路徑
                mould_board_path = self.get_mould_board_path(account_day_str, whole_file_name,
                                                             account_day_by_day)  # 模板ga位置，不知道有沒有.
                logging.info(f"Processing file: {mould_board_path}")
                if mould_board_path.is_file():  # 確認是否有檔案
                    # 叫pycharm安靜
                    # noinspection PyBroadException
                    try:
                        shutil.copy(mould_board_path, storage_path / whole_file_name)
                        log.append([client, account_day, "Success"])
                        break
                    except Exception as msg:  # 不知道例外處理方式，如果有檔案卻抓不出來，記錄下
                        logging.exception(f"Unexpected exception occurred: {msg}")
                        log.append([client, account_day, msg])
                        break
                else:
                    account_day_by_day -= 1
                    account_day -= timedelta(days=1)
                    if account_day_by_day == 0:
                        log.append([client, "本月無此模板", "Failed"])  # 如果日期歸零，代表沒模板
                        break
                    else:
                        account_day_str = account_day.strftime("%m%d")
        return log

    # 平日彙總
    def get_summary_path(self, account_day_str, whole_file_name, account_day_ymd, account_day_by_day):  # 取得ga上彙總放置路徑
        last_account_day_in_ga_base_account = self.get_last_account_day_in_ga_base_account()
        if account_day_by_day >= last_account_day_in_ga_base_account:
            summary_path = GA_BASE_ACCOUNT_PUBLIC_DIR / ("帳務日" + account_day_str) / "帳務" / (
                    whole_file_name + "-" + account_day_ymd) / "7.报表凭证"
        else:
            summary_path = GA_BASE_ACCOUNT_PUBLIC_DIR / (
                datetime.strptime(account_day_ymd, "%y%m%d").strftime("%Y.%#m")) / (
                                   "帳務日" + account_day_str) / "帳務" / (
                                   whole_file_name + "-" + account_day_ymd) / "7.报表凭证"
        return summary_path

    def move_summary(self, file_needed_list, log):
        storage_path = self.get_storage_path(folder_name="所需彙總")  # 儲存路徑
        if not storage_path.exists():
            self.creat_folder(folder_name="所需彙總")
        for client in file_needed_list:
            account_day_by_day = int(self.str_date_ymd[-2:])
            account_day_str = self.str_date_ymd[-4:]
            account_day_ymd = self.str_date_ymd
            account_day = datetime.strptime(self.str_date_ymd, "%y%m%d").date()
            while int(account_day_by_day) > 0:
                #  完整路徑
                summary_path = self.get_summary_path(account_day_str, client, account_day_ymd,
                                                     account_day_by_day)  # 模板ga位置，不知道有沒有.
                for file in summary_path.glob("*"):
                    if client.upper() in file.stem.upper() and "汇总" in file.stem and file.stat().st_size > 5000 and "~$" not in file.stem and "._" not in file.stem:
                        try:
                            shutil.copy(file, storage_path / (client + "-汇总" + file.suffix))
                            log.append([client, account_day, "Success"])
                            break
                        except Exception as msg:  # 不知道例外處理方式，如果有檔案卻抓不出來，記錄下
                            logging.exception(f"Unexpected exception occurred: {msg}")
                            log.append([client, account_day, msg])
                            break
                if (storage_path / (client + "-汇总.xlsx")).exists():
                    break

                account_day_by_day -= 1
                account_day -= timedelta(days=1)
                if account_day_by_day == 0:
                    log.append([client, "本月無此彙總", "Failed"])  # 如果日期歸零，代表沒模板
                    break
                else:
                    account_day_str = account_day.strftime("%m%d")
                    account_day_ymd = account_day.strftime("%y%m%d")
        return log

    # 平日帳務
    def move_today_account(self, file_needed_list, log):
        """
        設定帳務儲存路徑為桌面帳務資料夾
        若沒有此資料夾創建一個
        取的沒有年份的日期字串
        用get_account_path取的GA帳務路徑
        取得大寫名稱的需求客戶名稱
        取得GA帳務放置路徑底下帳務的迭代器
        用此迭代器取得過濾後剩下的路徑，過濾方式為資料夾名稱轉為大寫後，用-分列的第一個字，帳務資料夾為客戶名稱-日期，最後抓取這些到桌面

        :param file_needed_list:需求的客戶名單
        :param log:
        :return:
        """
        storage_path = self.get_storage_path(folder_name="帳務")  # 儲存路徑
        if not storage_path.exists():
            self.creat_folder(folder_name="帳務")
        account_day_str = self.str_date_ymd[-4:]
        account_path = self.get_account_path(account_day_str)
        account_needed_list_upper = [account.upper() for account in file_needed_list]
        account_iterator = account_path.glob("*")
        account_filter = filter(lambda folder:
                                "-".join(folder.name.upper().split("-")[:-1]) in account_needed_list_upper and
                                folder.is_dir(),
                                account_iterator)
        for account_folder in account_filter:
            folder_name_split: list = account_folder.name.split("-")
            try:
                shutil.copytree(account_folder, storage_path / account_folder.name)
                log.append(["-".join(folder_name_split[:-1]), folder_name_split[-1], "複製成功"])
            except Exception as msg:  # 不知道例外處理方式，如果有檔案卻抓不出來，記錄下
                logging.exception(f"Unexpected exception occurred: {msg}")
                log.append(["-".join(folder_name_split[:-1]), folder_name_split[-1], "複製失敗"])
        account_needed_upper_set: set = {account.upper() + "-" + self.str_date_ymd for account in file_needed_list}
        account_in_storage_path_set: set = {folder.name.upper() for folder in storage_path.glob("*")}
        short_account = account_needed_upper_set.difference(account_in_storage_path_set)
        for account_folder in short_account:
            folder_name_split: list = account_folder.split("-")
            log.append(["-".join(folder_name_split[:-1]), folder_name_split[-1], "沒找到此帳務"])

        return log

    # 問題整合
    @staticmethod
    def get_question_file_name(upper_database_dict, clients_name):  # 取的問題回復檔名
        question_file_name = upper_database_dict[str(clients_name).upper()] + "_" + clients_name + "問題及回覆"
        return question_file_name

    def get_question_file_path(self, whole_file_name):  # 取得問題回復檔案放置路徑，月份需要改
        question_file_path = (GA_BASE_ACCOUNT_PUBLIC_DIR / "06-02-04-問題整合" /
                              ("20" + self.str_date_ymd[:2] + "-" + str(int(self.str_date_ymd[2:4]))) /
                              whole_file_name)
        return question_file_path

    def move_question(self, file_needed_list, log):
        upper_database_dict = self.database.site_name_number_upper_dict
        storage_path = self.get_storage_path(folder_name="問題整合")
        if not storage_path.exists():
            self.creat_folder(folder_name="問題整合")
        for client in file_needed_list:
            question_file_name = self.get_question_file_name(upper_database_dict, client)
            whole_file_name = self.get_whole_file_name(question_file_name)  # 完整檔案包含副檔案名稱
            question_file_path = self.get_question_file_path(whole_file_name)
            logging.info(f"Processing file: {question_file_path}")
            if question_file_path.is_file():  # 確認是否有檔案
                # 叫pycharm安靜
                # noinspection PyBroadException
                try:
                    shutil.copy(question_file_path, storage_path / whole_file_name)
                    log.append([client, "success"])
                except Exception as msg:  # 不知道例外處理方式，如果有檔案卻抓不出來，記錄下
                    logging.exception(f"Unexpected exception occurred: {msg}")
                    log.append([client, msg])
            else:
                log.append([client, "找不到問題"])
        return log

    # 執行動作

    def get_file(self, person, action):
        folder_map: dict = {"over_summary": "關帳彙總",
                            "over_mould_board": "關帳模板",
                            "today_account": "帳務",
                            "summary": "所需彙總",
                            "mould_board": "當日模板",
                            "question": "問題整合"}

        folder_name = folder_map[action]
        fun_var = f"move_{action}"
        log_file_name = f"#get_{action}_log.xlsx"
        log = []
        if person == "自定清單":
            requirements_file_path = "自定檔案.xlsx"
            try:
                file_needed_list = self.get_file_needed_list(requirements_file_path)
            except Exception: # noqa
                file_needed_list = []
        else:
            file_needed_list = self.distribution.distribute_for_person_dict[person]
        log = getattr(self, fun_var)(file_needed_list, log)
        self.add_log(folder_name, log_file_name, log)
