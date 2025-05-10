import tkinter as tk
import time
import warnings
import numpy as np
from tkinter import messagebox
import pandas as pd
from pathlib import Path
import threading
from pandas import ExcelFile
import multiprocessing
import tkinter.messagebox as mbox
from concurrent.futures import ProcessPoolExecutor
from style import SummaryStyle
from bs4 import BeautifulSoup

CERTIFICATE1_KEYWORDS = []
CERTIFICATE2_KEYWORDS = []
CERTIFICATE3_KEYWORDS = []
CERTIFICATE4_KEYWORDS = []


GA_CERTIFICATE = Path("Cloud_Drive_path")  # 雲端憑證路徑
DESKTOP = Path.home() / 'Desktop'  # 桌面路徑
DESKTOP_MOULD = DESKTOP / "模板"  # 客戶檔案存放路徑
DESKTOP_ACCOUNT = DESKTOP / "帳務"  # 桌面帳務路徑

CLIENT_DATA = Path("Cloud_Drive_path / client_assignments.xlsx")  # 客戶工作分配表

class SummeryTable:
    def __init__(self, date_range):
        self.file_path = DESKTOP / "總表.xlsx"
        self.date_range = date_range
        self.start_date = pd.to_datetime(self.date_range.split("-")[0], format="%y%m%d")
        self.end_date = pd.to_datetime(self.date_range.split("-")[-1], format="%y%m%d")
        self.difference_columns = ["人數差額", "筆數差額", "金額差额"]
        self.one_site_columns = ["客戶", "預設", "序号", "日期", "憑證種類", "人数", "笔数", "金额", "币别",
                                 "路徑", "憑證人數",
                                 "憑證筆數", "憑證總額", "人數差額", "筆數差額", "金額差额", "檢查"]
        self.total_site_columns = ["客戶", "預設", "序号", "日期", "憑證種類", "人数", "笔数", "金额", "币别",
                                   "子客戶", "路徑", "憑證人數",
                                   "憑證筆數", "憑證總額", "人數差額", "筆數差額", "金額差额", "檢查"]
        self.certificate_columns = ["憑證人數", "憑證筆數", "憑證總額"]
        self.path_column = "路徑"
        self.check_column = "檢查"
        self.sheet1_name = "單客戶"
        self.sheet2_name = "多客戶"
        self.sheet3_name = "失敗檔案"
        self.data_sheet_list = [self.sheet1_name, self.sheet2_name]

    def get_client_system_name(self, client_name):
        """
        讀取工作分配表上的客戶清單
        :param client_name:
        :return:
        """
        # noinspection PyBroadException
        try:
            str_date_ymd = self.date_range.split("-")[-1]
            sheet_name = str_date_ymd[0:2] + "-" + str(int(str_date_ymd[2:4]))
            full_df = pd.read_excel(CLIENT_DATA, sheet_name=sheet_name)
            # 一般盤口
            df = full_df[["客戶名稱", "系统"]].copy()
            # 特殊客戶1
            df2 = full_df[["客戶名稱.1", "系统.1"]].copy()
            df2.columns = ["客戶名稱", "系统"]
            # 特殊客戶12
            df3 = full_df[["客戶名稱.2", "系统類"]].copy()
            df3.columns = ["客戶名稱", "系统"]

            df["客戶名稱"] = df["客戶名稱"].str.upper()
            df2["盤口名稱"] = df2["客戶名稱"].str.upper()
            df3["客戶名稱"] = df3["客戶名稱"].str.upper()
            total_df = pd.concat([df, df2, df3])
            system_name = total_df[total_df["客戶名稱"] == client_name.upper()]["系统"].values[0]

        except Exception:
            return None
        return system_name

    @staticmethod
    def get_update_file(file_list: list):
        """
        number_for_n為創建一個將-n轉換為數字大小的列表，預設空
        將list轉為小寫，因為有人寫大寫有人寫小寫
        沒有-n檔案，算0
        -n算是1
        -n1算是2
        因為for有順序性，因此append進去number_for_n也會依照順序
        接下來判別數字，列表中最大值代表最新的檔案
        數字不能等於0，代表沒有-n的名字有兩個，因為傳入的len(list)必定大於1
        數字大於1，則計算此數字在列表中有幾個，1個才回傳
        :param file_list:後臺數據中符合序號、日期、關鍵字過濾後的檔案路徑list(包含-n)，len(list)必定大於1
        :return:回傳-n最新的檔案路徑，又或是檔案有誤時回傳None
        """
        number_for_n = []
        lower_file_list = [file.stem.lower() for file in file_list]  # -n有人大寫有人小寫
        for file in lower_file_list:
            try:
                number = int(file.split("-n")[-1]) + 1
            except ValueError:
                if (file.split("-n")[-1]) == "":
                    number = 1
                else:
                    number = 0
            number_for_n.append(number)
        max_number = max(number_for_n)  # 找最大的數字
        if max_number == 0:  # 數字不能等於0，代表沒有-n的名字有兩個
            return None
        elif max_number > 0:
            if number_for_n.count(max_number) != 1:  # 最大值也不能是兩個以上，代表有重複的-n數字
                return None
            else:
                index = number_for_n.index(max_number)
                return file_list[index]
        return None

    def looking_for_file(self, index, str_day_ymd, subject, filter_file_list):
        """
        尋找路徑檔案
        先過濾掉暫存檔案跟符合序號、日期的檔案，以及格式，在過濾科目
        剩餘的list，在過濾科目
        如果list為0，代表沒有相應檔案
        list為1，代表就是此檔案，故回傳
        list大於1，代表有多個檔案需判別，故執行get_update_file並回傳
        :param index: dataframe中序號
        :param str_day_ymd: dataframe日期
        :param subject: dataframe科目
        :param filter_file_list: 6資料夾中全部的檔案
        :return: 回傳最新的檔案路徑
        """
        filter_file = filter(lambda file: (index == file.stem.split("-")[0]
                                           and
                                           str_day_ymd in file.stem
                                           and "~" not in file.stem
                                           and ((file.stem.count("-") == 2 and file.stem.lower().count("n") == 0)
                                                or (file.stem.count("-") == 3 and file.stem.lower().count("n") == 1))),
                             filter_file_list)

        if subject == "CERTIFICATE1":
            filter_file = filter(lambda file: (any(key in file.stem for key in CERTIFICATE1_KEYWORDS)),
                                 filter_file)
        elif subject == "CERTIFICATE2":
            filter_file = filter(lambda file: (any(key in file.stem for key in CERTIFICATE2_KEYWORDS)),
                                 filter_file)
        elif subject == "CERTIFICATE3":
            pass
        elif subject == "CERTIFICATE4":
            filter_file = filter(lambda file: (any(key in file.stem for key in CERTIFICATE4_KEYWORDS)),
                                 filter_file)

        filter_file_list = list(filter_file)
        if len(filter_file_list) == 0:  # 代表沒找到
            return None
        elif len(filter_file_list) == 1:  # 代表僅有一個
            return filter_file_list[0]
        else:
            return self.get_update_file(filter_file_list)  # 代表多個，要找到正確的

    def get_file_path(self, row):
        """
        file_path_parent為檔案應該在的路徑，此處設定為憑證區
        file_list取得file_path_parent底下所有檔案
        使用self.looking_for_file取得唯一檔案
        :param row:傳入dataframe整行資料，用於apply
        :return:
        """
        index = str(row["序号"])
        account_day_ymd = row["日期"].strftime("%y%m%d")
        account_day_full = row["日期"].strftime("%Y%m%d")
        subject = row["憑證種類"]
        file_path_parent = (GA_CERTIFICATE / (account_day_full[:4] + "-" + str(int(account_day_full[4:6]))) /
                            row["客戶"] / "6.数据")
        file_list = [file for file in file_path_parent.glob("*")]
        file = self.looking_for_file(index, account_day_ymd, subject, file_list)
        return file

    @staticmethod
    def get_certificate_type(row, client_system_name):
        """
        依照客戶系統別分配對應的憑證Class
        :param row:傳入dataframe整行資料
        :param client_system_name: 傳入客戶系統別
        :return: 回傳已指定class的certificate或是None
        """
        mapping = {
            "System1": CertificateForSystem1,
            "System2": CertificateForSystem2,
            "System3": CertificateForSystem3,
            "System4": CertificateForSystem4,
            "System5": CertificateForSystem5,
            "System6": CertificateForSystem6,
            "System7": CertificateForSystem7,
            "System8": CertificateForSystem8,
            "System9": CertificateForSystem9,
            "System10": CertificateForSystem10,
            "System11": CertificateForSystem11,
            "System12": CertificateForSystem12,
            "System13": CertificateForSystem13,
            "System14": CertificateForSystem14
        }

        if client_system_name in mapping:
            return mapping[client_system_name](row["路徑"], row["憑證種類"], row["客戶"])
        else:
            return None

    def get_data_for_certificate(self, row, total_site, total_site_series) -> pd.Series:
        """
        用self.get_client_system_name取得站點對應的系統別

        若系統別存在使用get_certificate_type取得已分配class的certificate，否則回傳人數、筆數、、總額空值
        若self.get_certificate_type沒有取得certificate也回傳人數、筆數、總額空值，因為可能新站或是新系統尚未建立

        若有憑證路徑，使用certificate.get_data_for_certificate()取的憑證人數、筆數、總額，否則回傳空值
        :param total_site_series: 依照序號對應的站點series
        :param total_site: 是否為總站，True或False
        :param row: 傳入dataframe整行資料
        :return:回傳人數筆數總額的Series
        """
        if not total_site:
            client_system_name = self.get_client_system_name(row["客戶"])
        else:
            client_system_name = self.get_client_system_name(total_site_series[row["序号"]])
        empty_series = pd.Series([None, None, None], index=self.certificate_columns)
        if client_system_name:
            certificate = self.get_certificate_type(row, client_system_name)
        else:
            return empty_series

        if certificate is None:
            return empty_series

        if certificate.file_path:
            sum_of_amount, number_of_people, number_of_data = certificate.get_data_for_certificate()
            data_series = pd.Series([number_of_people, number_of_data, sum_of_amount], index=self.certificate_columns)
        else:
            return empty_series

        return data_series

    def get_df_top_up_withdraw(self, mould) -> (pd.DataFrame, list):
        """
        使用模板取得充值的dataframe
        新增盤口名的column
        篩選日期，重新排序欄位
        若dataframe不為空，則刪除list此檔案名稱
        使用apply取得檔案對應路徑在取得該路徑檔案中的人數、筆數、總額
        若是空的，則回傳空dataframe
        :param mould: 模板物件
        :return:充值提現的dataframe、資料夾中檔案的剩餘list
        """
        df_top_up_withdraw = mould.df_mould_top_up_withdraw
        df_top_up_withdraw["客戶"] = mould.name
        if not pd.api.types.is_datetime64_any_dtype(df_top_up_withdraw["日期"]):  # 低概率模版本身会有误，重新限定日期属性
            df_top_up_withdraw["日期"] = pd.to_datetime(df_top_up_withdraw["日期"])
        df_top_up_withdraw = df_top_up_withdraw[
            (self.start_date <= df_top_up_withdraw["日期"]) & (df_top_up_withdraw["日期"] <= self.end_date)]
        df_top_up_withdraw = df_top_up_withdraw[
            ["客戶", "序号", "日期", "憑證種類", "人数", "笔数", "金额", "币别"]]  # 重新排序
        if not df_top_up_withdraw.empty:
            df_top_up_withdraw[self.path_column] = df_top_up_withdraw.apply(self.get_file_path, axis=1)
            df_top_up_withdraw[self.certificate_columns] = df_top_up_withdraw.apply(
                self.get_data_for_certificate, args=(mould.total_site, mould.total_site_series), axis=1)
        else:
            return pd.DataFrame(
                columns=["客戶", "序号", "日期", "憑證種類", "人数", "笔数", "金额", "币别", "路徑", "憑證人數",
                         "憑證筆數", "憑證總額"]), mould.full_name
        return df_top_up_withdraw, None

    @staticmethod
    def difference_round(df_data, difference_column):
        """
        差異欄位的四捨五入
        :param df_data:
        :param difference_column:
        :return:
        """
        if not df_data[difference_column].isna().all():
            df_columns_not_na = df_data[difference_column].notna()
            df_data.loc[df_columns_not_na, difference_column] = (df_data.loc[df_columns_not_na, difference_column].
                                                                 astype(float).round(4))
        return df_data

    def get_summary_difference(self, df_data):
        df_data["純數值人數"] = pd.to_numeric(df_data["人数"], errors='coerce')
        df_data["純數值笔數"] = pd.to_numeric(df_data["笔数"], errors='coerce')
        df_data["純數值金額"] = pd.to_numeric(df_data["金额"], errors='coerce')
        df_data["人數差額"] = (df_data["純數值人數"] - df_data["憑證人數"])
        df_data["筆數差額"] = (df_data["純數值笔數"] - df_data["憑證筆數"])
        df_data["金額差额"] = (df_data["純數值金額"] - df_data["憑證總額"])
        for column in self.difference_columns:
            df_data = self.difference_round(df_data, column)
        df_data = df_data.drop(columns=["純數值人數", "純數值笔數", "純數值金額"])
        df_data[self.check_column] = (
                (df_data["人數差額"] == 0) & (df_data["筆數差額"] == 0) & (df_data["金額差额"] == 0))
        return df_data

    def df_to_excel(self, df_all_data, df_all_data_total_site, df_remaining_file):
        with pd.ExcelWriter(self.file_path) as writer:
            df_all_data.to_excel(writer, self.sheet1_name, index=False)
            df_all_data_total_site.to_excel(writer, self.sheet2_name, index=False)
            df_remaining_file.to_excel(writer, self.sheet3_name, index=False)

    def process(self, file_path):
        try:
            mould = ClientMould(file_path)
            df_data, remaining_file = self.get_df_top_up_withdraw(mould)
            return df_data, remaining_file, mould.total_site, mould.total_site_series
        except Exception:
            return None, file_path.name, None, None

    def get_summary_table(self):
        df_all_data = pd.DataFrame(
            columns=self.one_site_columns)
        df_all_data_total_site = pd.DataFrame(
            columns=self.total_site_columns)
        with ProcessPoolExecutor(max_workers=multiprocessing.cpu_count()) as executor:
            futures = []
            for file_path in DESKTOP_MOULD.glob("*"):
                if file_path.is_file():
                    futures.append(executor.submit(self.process, file_path))

            df_all_data, df_all_data_total_site, remaining_file_list = self.get_result(futures, df_all_data,
                                                                                       df_all_data_total_site)
        df_remaining_file = pd.DataFrame(remaining_file_list, columns=[self.sheet3_name])
        self.df_to_excel(df_all_data, df_all_data_total_site, df_remaining_file)

    def update_summary_table(self):
        df_all_data, df_all_data_total_site, df_remaining_file = self.get_summary_data()
        df_all_data_need_update = df_all_data[(df_all_data[self.check_column] == False)].copy()  # noqa :E712
        df_all_data_total_site_need_update = df_all_data_total_site[
            (df_all_data_total_site[self.check_column] == False)].copy()

        total_site_series = pd.Series(data=df_all_data_total_site_need_update['客戶'].values,
                                      index=df_all_data_total_site_need_update['序号'])
        df_list = [df_all_data, df_all_data_total_site]
        df_update_list = [df_all_data_need_update, df_all_data_total_site_need_update]

        total_site_info = [(False, None), (True, total_site_series)]
        for i, (df, df_need_update, (site, series)) in enumerate(zip(df_list, df_update_list, total_site_info)):
            if not df.empty:
                df_list[i] = self.data_update(df, df_need_update, site, series)

        self.df_to_excel(df_list[0], df_list[1], df_remaining_file)

    def data_update(self, df, df_need_update, total_site, total_site_series):
        df_need_update[self.path_column] = df_need_update.apply(self.get_file_path, axis=1)
        df_need_update[self.certificate_columns] = df_need_update.apply(
            self.get_data_for_certificate, args=(total_site, total_site_series),
            axis=1)
        df.loc[(df["檢查"] == False), [self.path_column]] = df_need_update[self.path_column]  # noqa :E712
        df.loc[(df["檢查"] == False), self.certificate_columns] = df_need_update[
            self.certificate_columns]  # noqa :E712
        return df

    def update_fail(self):
        df_all_data, df_all_data_total_site, df_remaining_file = self.get_summary_data()
        remaining_file_list = df_remaining_file[self.sheet3_name].tolist()

        with ProcessPoolExecutor(max_workers=multiprocessing.cpu_count()) as executor:
            futures = []
            for file_name in remaining_file_list.copy():
                file_path = DESKTOP_MOULD / file_name
                if file_path.is_file():
                    futures.append(executor.submit(self.process, file_path))
            df_all_data, df_all_data_total_site, remaining_file_list = self.get_result(futures, df_all_data,
                                                                                       df_all_data_total_site)

        df_remaining_file = pd.DataFrame(remaining_file_list, columns=[self.sheet3_name])
        self.df_to_excel(df_all_data, df_all_data_total_site, df_remaining_file)

    def get_summary_data(self):
        df_all_data = pd.read_excel(self.file_path, sheet_name=self.sheet1_name)
        df_all_data = self.get_summary_difference(df_all_data)  # 因为style写入公式，因此摆到前面 原先在208
        df_all_data_total_site = pd.read_excel(self.file_path, sheet_name=self.sheet2_name)
        df_all_data_total_site = self.get_summary_difference(df_all_data_total_site)
        df_remaining_file = pd.read_excel(self.file_path, sheet_name=self.sheet3_name)
        return df_all_data, df_all_data_total_site, df_remaining_file

    @staticmethod
    def get_result(futures, df_all_data, df_all_data_total_site):
        remaining_file_list = []
        for future in futures:
            df_data, remaining_file, total_site, total_site_series = future.result()
            if df_data is not None:
                if total_site:
                    df_data["站点"] = df_data["序号"].map(total_site_series)
                    df_all_data_total_site = pd.concat([df_all_data_total_site, df_data])
                else:
                    df_all_data = pd.concat([df_all_data, df_data])
            if remaining_file is not None:
                remaining_file_list.append(remaining_file)
        return df_all_data, df_all_data_total_site, remaining_file_list


class SummeryTableByDeskTop(SummeryTable):
    def get_file_path(self, row):
        index = str(row["序号"])
        account_day_ymd = row["日期"].strftime("%y%m%d")
        subject = row["憑證種類"]
        file_path_parent = DESKTOP_ACCOUNT / (row["客戶"] + "-" + account_day_ymd) / "6.数据"
        file_list = [file for file in file_path_parent.glob("*")]
        file = self.looking_for_file(index, account_day_ymd, subject, file_list)
        return file


class ClientMould:
    def __init__(self, file_path):
        self.file_path = file_path
        self.full_name = file_path.name
        self.name = file_path.stem
        self.total_site = False
        self.total_site_series = None
        self.dtype = {"序号": int, "憑證種類": str, "人数": float, "笔数": float, "金额": float, "币别": str}
        self.usecols = ["序号", "日期", "憑證種類", "人数", "笔数", "金额", "币别"]
        self.df_mould_top_up_withdraw = self.get_df_mould_top_up_withdraw()

    def get_df_mould_top_up_withdraw(self):
        # noinspection PyTypeChecker
        try:
            df_top_up_withdraw = (pd.read_excel(self.file_path,
                                                sheet_name="憑證數據",
                                                header=1,
                                                parse_dates=["日期"])).dropna(subset="金额")
        except Exception as msg:
            print(self.name + "模版有误" + str(msg))
            return pd.DataFrame(
                columns=["客戶", "序号", "日期", "三级科目", "人数", "笔数", "金额", "币别", "路徑",
                         "憑證人數",
                         "憑證筆數", "憑證總額"])

        if "子客戶" in df_top_up_withdraw.columns:
            self.total_site = True
            self.total_site_series = pd.Series(data=df_top_up_withdraw['子客戶'].values, index=df_top_up_withdraw['序号'])

        try:
            df_top_up_withdraw = df_top_up_withdraw[self.usecols]
            for columns, dtype in self.dtype.items():
                df_top_up_withdraw[columns] = df_top_up_withdraw[columns].astype(dtype)

        except ValueError as msg:  # todo 这边应该要改成所有错误，不然如果模版有加密会错误
            df_top_up_withdraw = df_top_up_withdraw[self.usecols]
            print(f"{self.name}有奇怪資料，限定屬性失敗 {msg}")

        return df_top_up_withdraw

class CertificateForSystem1:
    def __init__(self, file_path, subject, client_name):
        self.file_path = file_path
        self.subject = subject
        self.client_name = client_name
        (self.usecols,
         self.dtype,
         ) = self.get_read_parameter()
        self.subset = self.usecols
        if self.usecols:
            self.order_number = self.usecols[0]
            self.number_of_members = self.usecols[1]
            self.sum_columns = self.usecols[2]
        else:
            self.order_number = None
            self.number_of_members = None
            self.sum_columns = None

    def get_data_for_certificate(self):
        print(self.file_path)
        try:
            excel_file = ExcelFile(self.file_path)
            sum_of_amount, number_of_people, number_of_data = self.data_for_file_excel(excel_file)
        except ValueError as msg:
            if str(msg) == "Excel file format cannot be determined, you must specify an engine manually.":
                sum_of_amount, number_of_people, number_of_data = self.data_for_file_csv()
                if (sum_of_amount, number_of_people, number_of_data) == (None, None, None):
                    sum_of_amount, number_of_people, number_of_data = self.data_for_file_xml()
            else:
                return None, None, None
        except Exception as msg:
            print("檔案無法以excel開啟" + str(msg))
            return None, None, None

        return sum_of_amount, number_of_people, number_of_data

    def get_read_parameter(self):
        if self.subject == "CERTIFICATE2" or self.subject == "CERTIFICATE1":
            usecols = ["订单号", "会员账号", "订单金额", "订单状态"]
            dtype = {"订单号": str, "会员账号": str, "金额": str, "订单状态": str}

        elif self.subject == "CERTIFICATE4":
            usecols = ["订单号", "会员账号", "提现金额", "订单状态"]
            dtype = {"订单号": str, "会员账号": str, "金额": str, "订单状态": str}
        else:
            usecols, dtype = None, None
        return usecols, dtype

    def df_data_for_file_filter_condition(self, df_data_for_file):
        if self.subject == "CERTIFICATE4":
            df_data_for_file = df_data_for_file[(df_data_for_file[self.usecols[3]] == "成功") |
                                                (df_data_for_file[self.usecols[3]] == "手動完成")
                                                ]
        else:
            df_data_for_file = df_data_for_file[df_data_for_file[self.usecols[3]] == "成功"]
        return df_data_for_file

    def data_for_file_xml(self):
        number_of_people, number_of_data, sum_of_amount = None, None, None
        try:
            warnings.filterwarnings("ignore", category=UserWarning)
            soup = BeautifulSoup(open(self.file_path, encoding="utf-8"), features="lxml")
            worksheet_list = [worksheet.get('ss:name') for worksheet in soup.find_all('worksheet')]
            for sheet in worksheet_list:
                worksheet = soup.find('worksheet', attrs={'ss:name': f"{sheet}"})
                try:
                    rows = worksheet.find_all('row')
                    columns = [cell.get_text() for cell in rows[0]]
                    data_list = [[cell.get_text() for cell in row][:len(columns)] for row in rows[1:]]
                    full_df = pd.DataFrame(data_list, columns=columns)

                    df_data_for_file = full_df[self.usecols]
                    df_data_for_file = df_data_for_file.replace("", np.nan)
                    df_data_for_file = df_data_for_file.dropna(subset=self.subset)

                    for column, dtype in self.dtype.items():
                        df_data_for_file[column] = df_data_for_file[column].astype(dtype)
                    sum_of_amount, number_of_people, number_of_data = self.get_statistical_data(df_data_for_file,
                                                                                                self.subset,
                                                                                                self.sum_columns,
                                                                                                self.number_of_members,
                                                                                                self.order_number)
                    return sum_of_amount, number_of_people, number_of_data
                except Exception:
                    continue
            return sum_of_amount, number_of_people, number_of_data
        except Exception as msg:
            print(msg)
            return sum_of_amount, number_of_people, number_of_data

    def data_for_file_csv(self):
        number_of_people, number_of_data, sum_of_amount = None, None, None
        encoding_list = ["GBK", "utf-8"]
        for encoding in encoding_list:
            # noinspection PyBroadException
            try:
                df_data_for_file = (pd.read_csv(self.file_path,
                                                usecols=self.usecols,
                                                encoding=encoding,
                                                dtype=self.dtype,
                                                na_values="",
                                                keep_default_na=False))
                if df_data_for_file.empty:
                    continue
                sum_of_amount, number_of_people, number_of_data = self.get_statistical_data(df_data_for_file,
                                                                                            self.subset,
                                                                                            self.sum_columns,
                                                                                            self.number_of_members,
                                                                                            self.order_number)
                break
            except UnicodeDecodeError:  # UnicodeDecodeError
                continue
            except Exception as msg:
                print(self.file_path)
                print(msg)
                continue

        return sum_of_amount, number_of_people, number_of_data

    def read_excel(self, sheet):
        # noinspection PyTypeChecker
        return (pd.read_excel(self.file_path, sheet_name=sheet,
                              usecols=self.usecols, dtype=self.dtype,
                              na_values="", keep_default_na=False))

    def data_for_file_excel(self, excel_file):
        number_of_people, number_of_data, sum_of_amount = None, None, None

        for sheet in excel_file.sheet_names:
            # noinspection PyBroadException
            try:
                # noinspection PyTypeChecker
                df_data_for_file = self.read_excel(sheet)

                if df_data_for_file.empty:
                    continue

                sum_of_amount, number_of_people, number_of_data = self.get_statistical_data(df_data_for_file,
                                                                                            self.subset,
                                                                                            self.sum_columns,
                                                                                            self.number_of_members,
                                                                                            self.order_number)

                break
            except Exception as msg:  # ValueError or KeyError
                print(self.file_path)
                print(msg)
                continue
        return sum_of_amount, number_of_people, number_of_data

    def get_statistical_data(self, df_data_for_file, subset, sum_columns, number_of_members, order_number):
        df_data_for_file.dropna(subset=subset, inplace=True)
        df_data_for_file[self.usecols[2]] = df_data_for_file[self.usecols[2]].astype(float)
        df_data_for_file = self.df_data_for_file_filter_condition(df_data_for_file)
        sum_of_amount = round(df_data_for_file[sum_columns].sum(), 4)
        number_of_people = df_data_for_file[number_of_members].nunique()
        number_of_data = df_data_for_file[order_number].count()
        return sum_of_amount, number_of_people, number_of_data


class CertificateForSystem2(CertificateForSystem1):
    def __init__(self, file_path, subject, client_name):
        super().__init__(file_path, subject, client_name)
        self.header = self.get_excel_header()

    def get_excel_header(self):
        certificate1_client_list = ["Client1", "Client2", "Client3", "Client4", "Client5",
                               "Client6"]
        certificate2_client_list = ["Client1", "Client2"]

        if self.subject == "CERTIFICATE1" and self.client_name in certificate1_client_list:
            header = 1
        elif self.subject == "CERTIFICATE2" and self.client_name in certificate2_client_list:
            header = 1
        else:
            header = 0
        return header

    def get_read_parameter(self):
        no_type_by_certificate1_group = ["Group1", "Group2", "Group3"]
        no_type_by_certificate2_group = ["Group1", "Group2", "Group3", "Group4",
                                   "Group5",
                                   "Group6", "Group7", "Group8",
                                   "Group9"]
        no_type_by_all_group = ["AllGroup1", "AllGroup2", "AllGroup3"]
        special_parameter_group = ["SpecialGroup1", "SpecialGroup2", "SpecialGroup3", "SpecialGroup4", 
                                   "SpecialGroup5", "SpecialGroup6",
                                   "SpecialGroup7",
                                   "SpecialGroup8",
                                   "SpecialGroup9"]
        special2_parameter_group = ["Special2Group1"]
        if self.subject == "CERTIFICATE1" or self.subject == "CERTIFICATE2":
            if self.client_name in no_type_by_certificate2_group:
                (usecols, dtype) = self.no_type_by_online_parameter()
            elif self.client_name in no_type_by_certificate1_group:
                (usecols, dtype) = self.no_type_by_company_parameter()
            elif self.client_name in no_type_by_all_group:
                (usecols, dtype) = self.no_type_by_all_parameter()
            elif self.client_name in special_parameter_group:
                (usecols, dtype) = self.special_parameter()
            elif self.client_name in special2_parameter_group:
                (usecols, dtype) = self.special2_parameter()
            else:
                (usecols, dtype) = self.default_parameter()
        else:
            usecols = ["编号", "会员账号", "资讯", "狀態"]
            dtype = {"编号": str, "会员账号": str, "资讯": str, "狀態": str}
        return usecols, dtype

    def default_parameter(self):  # 導出的加密貨幣與在線支付
        if self.subject == "CERTIFICATE1":
            usecols = ["订单号", "会员账号", "金额", "状态"]
            dtype = {"订单号": str, "会员账号": str, "金额": str, "状态": str}

        elif self.subject == "CERTIFICATE2":
            usecols = ["订单号", "会员账号", "金额", "状态"]
            dtype = {"订单号": str, "会员账号": str, "金额": str, "状态": str}

        else:
            usecols, dtype = None, None
        return usecols, dtype

    def no_type_by_online_parameter(self):
        if self.subject == "CERTIFICATE1":
            usecols = ["订单号", "会员账号", "金额", "状态"]
            dtype = {"订单号": str, "会员账号": str, "金额": str, "状态": str}

        elif self.subject == "CERTIFICATE2":
            usecols = ["NO.", "会员帐号", "金额"]
            dtype = {"NO.": str, "会员帐号": str, "金额": str}

        else:
            usecols, dtype = None, None
        return usecols, dtype

    def no_type_by_company_parameter(self):
        if self.subject == "CERTIFICATE1":
            usecols = ["NO.", "会员帐号", "金额"]
            dtype = {"NO.": str, "会员帐号": str, "金额": str}

        elif self.subject == "CERTIFICATE2":
            usecols = ["订单号", "会员账号", "金额", "状态"]
            dtype = {"订单号": str, "会员账号": str, "金额": str, "状态": str}

        else:
            usecols, dtype = None, None
        return usecols, dtype

    def no_type_by_all_parameter(self):
        if self.subject == "CERTIFICATE1" or self.subject == "CERTIFICATE2":
            usecols = ["NO.", "会员帐号", "金额"]
            dtype = {"NO.": str, "会员帐号": str, "金额": str}

        else:
            usecols, dtype = None, None
        return usecols, dtype

    def special_parameter(self):
        if self.subject == "CERTIFICATE1":
            usecols = ["订单号", "会员帐号", "金额"]
            dtype = {"订单号": str, "会员帐号": str, "金额": str}

        elif self.subject == "CERTIFICATE2":
            usecols = ["NO.", "会员帐号", "金额"]
            dtype = {"NO.": str, "会员帐号": str, "金额": str}

        else:
            usecols, dtype = None, None
        return usecols, dtype

    def special2_parameter(self):  # 公司入款訂單號跟預設線上支付
        if self.subject == "CERTIFICATE1":
            usecols = ["订单号", "会员帐号", "金额"]
            dtype = {"订单号": str, "会员帐号": str, "金额": str}

        elif self.subject == "CERTIFICATE2":
            usecols = ["订单号", "会员账号", "金额", "状态"]
            dtype = {"订单号": str, "会员账号": str, "金额": str, "状态": str}

        else:
            usecols, dtype = None, None
        return usecols, dtype

    def read_excel(self, sheet):
        # noinspection PyTypeChecker
        return (pd.read_excel(self.file_path, sheet_name=sheet, header=self.header,
                              usecols=self.usecols, dtype=self.dtype,
                              na_values="", keep_default_na=False))

    def df_data_for_file_filter_condition(self, df_data_for_file):
        if len(self.usecols) == 4:
            if self.subject == "CERTIFICATE1":
                df_data_for_file = df_data_for_file[(df_data_for_file[self.usecols[3]] == "成功") |
                                                    (df_data_for_file[self.usecols[3]] == "手動成功")]
            elif self.subject == "CERTIFICATE2":
                if "状态" in df_data_for_file.columns:
                    df_data_for_file = df_data_for_file[(df_data_for_file[self.usecols[3]] == "成功") |
                                                        (df_data_for_file[self.usecols[3]] == "手動成功")]
            elif self.subject == "CERTIFICATE4":
                df_data_for_file = df_data_for_file[df_data_for_file[self.usecols[3]] == "确定"]
        return df_data_for_file

    def get_statistical_data(self, df_data_for_file, subset, sum_columns, number_of_members, order_number):
        if self.subject in ["CERTIFICATE1", "CERTIFICATE2"]:
            df_data_for_file[self.usecols[2]] = pd.to_numeric(df_data_for_file[self.usecols[2]].str.replace(",", ""))
        return super().get_statistical_data(df_data_for_file, subset, sum_columns, number_of_members, order_number)


class CertificateForSystem3(CertificateForSystem1):  # 万博
    def get_read_parameter(self):
        if self.subject == "CERTIFICATE1" or self.subject == "CERTIFICATE2":
            usecols = ["订单号", "会员账号", "金额", "订单状态"]
            dtype = {"订单号": str, "会员账号": str, "金额": str, "订单状态": str}

        elif self.subject == "CERTIFICATE4":
            usecols = ["订单号", "用户账号", "金额"]
            dtype = {"订单号": str, "用户账号": str, "金额": str}

        else:
            usecols, dtype = None, None
        return usecols, dtype

    def df_data_for_file_filter_condition(self, df_data_for_file):
        if self.subject != "CERTIFICATE4":
            df_data_for_file = df_data_for_file[df_data_for_file[self.usecols[3]] == "已成功"]
        return df_data_for_file


class CertificateForSystem4(CertificateForSystem1):
    def get_read_parameter(self):
        if self.subject == "CERTIFICATE1":
            usecols = ["订单号", "會員帳號", "金額", "狀態"]
            dtype = {"订单号": str, "會員帳號": str, "金額": str, "狀態": str}

        elif self.subject == "CERTIFICATE2":
            usecols = ["订单号", "會員帳號", "金額", "狀態"]
            dtype = {"订单号": str, "會員帳號": str, "金額": str, "狀態": str}

        elif self.subject == "CERTIFICATE4":
            usecols = ["單號", "會員帳號", "金額", "狀態"]
            dtype = {"單號": str, "會員帳號": str, "金額": str, "狀態": str}

        else:
            usecols, dtype = None, None
        return usecols, dtype

    def df_data_for_file_filter_condition(self, df_data_for_file):
        df_data_for_file = df_data_for_file[df_data_for_file[self.usecols[3]] == "已確認"]
        return df_data_for_file

    @staticmethod
    def get_company_rebuild(df_data_for_file):
        # noinspection PyTypeChecker
        df_company_rebuild = df_data_for_file.dropna(subset=["订单号", "會員帳號", "狀態"]).copy().reset_index(
            drop=True)
        se_data_for_file_by_amount = df_data_for_file[df_data_for_file["金額"].str.contains("金額", na=False)][
            "金額"].str.split('：').str.get(1).astype(float).reset_index(drop=True)
        df_company_rebuild["金額"] = se_data_for_file_by_amount
        return df_company_rebuild

    def get_statistical_data(self, df_data_for_file, subset, sum_columns, number_of_members, order_number):
        if self.subject == "CERTIFICATE1":
            df_data_for_file = self.get_company_rebuild(df_data_for_file)
        else:
            df_data_for_file.dropna(subset=subset, inplace=True)
            df_data_for_file[self.usecols[2]] = df_data_for_file[self.usecols[2]].astype(float)
        df_data_for_file = self.df_data_for_file_filter_condition(df_data_for_file)
        sum_of_amount = round(df_data_for_file[sum_columns].sum(), 4)
        number_of_people = df_data_for_file[number_of_members].nunique()
        number_of_data = df_data_for_file[order_number].count()
        return sum_of_amount, number_of_people, number_of_data


class CertificateForSystem5(CertificateForSystem1):
    def get_read_parameter(self):
        if self.subject == "CERTIFICATE1" or self.subject == "CERTIFICATE2":
            usecols = ["流水号", "用户名", "金额", "状态"]
            dtype = {"流水号": str, "用户名": str, "金额": str, "状态": str}

        elif self.subject == "CERTIFICATE4":
            usecols = ["订单号", "用户名", "金额", "状态"]
            dtype = {"订单号": str, "用户名": str, "金额": str, "状态": str}

        else:
            usecols, dtype = None, None
        return usecols, dtype

    def df_data_for_file_filter_condition(self, df_data_for_file):
        if self.subject == "CERTIFICATE1" or self.subject == "CERTIFICATE2":
            df_data_for_file = df_data_for_file[df_data_for_file[self.usecols[3]] == "已通过"]
        elif self.subject == "CERTIFICATE4":
            df_data_for_file = df_data_for_file[df_data_for_file[self.usecols[3]] == "成功"]
        return df_data_for_file


class CertificateForSystem6(CertificateForSystem1):
    def get_read_parameter(self):
        if self.subject in ["CERTIFICATE1", "CERTIFICATE2"]:
            usecols = ["订单号", "会员", "金额", "状态"]
            dtype = {"订单号": str, "会员": str, "金额": str, "状态": str}

        elif self.subject == "CERTIFICATE4":
            usecols = ["订单号", "会员", "金额", "状态"]
            dtype = {"订单号": str, "会员": str, "金额": str, "状态": str}

        else:
            usecols, dtype = None, None
        return usecols, dtype

    def df_data_for_file_filter_condition(self, df_data_for_file):
        if self.subject == "CERTIFICATE1":
            df_data_for_file = df_data_for_file[df_data_for_file[self.usecols[3]] == "已成功"]
        if self.subject == "CERTIFICATE2":
            df_data_for_file = df_data_for_file[df_data_for_file[self.usecols[3]] == "成功"]
        elif self.subject == "CERTIFICATE4":
            df_data_for_file = df_data_for_file[df_data_for_file[self.usecols[3]] == "完成"]
        return df_data_for_file

    def get_statistical_data(self, df_data_for_file, subset, sum_columns, number_of_members, order_number):
        if self.subject == "CERTIFICATE4":

            self.usecols = ["订单号", "经销商", "金额", "状态"]
            self.dtype = {"订单号": str, "经销商": str, "金额": str, "状态": str}
            self.sum_columns = self.usecols[2]
            df_data_for_file_dealer = pd.DataFrame(columns=[self.usecols])
            excel_file = ExcelFile(self.file_path)
            for sheet in excel_file.sheet_names:
                try:
                    df_data_for_file_dealer = self.read_excel(sheet)
                    df_data_for_file_dealer = df_data_for_file_dealer.rename(columns={self.usecols[1]: "会员"})
                    break
                except Exception:
                    continue
            df_data_for_file = pd.concat([df_data_for_file, df_data_for_file_dealer])
            df_data_for_file.reset_index(drop=True, inplace=True)
        sum_of_amount, number_of_people, number_of_data = super().get_statistical_data(df_data_for_file, subset,
                                                                                       sum_columns, number_of_members,
                                                                                       order_number)
        return sum_of_amount, number_of_people, number_of_data


class CertificateForSystem7(CertificateForSystem1):
    def get_read_parameter(self):
        if self.subject in ["CERTIFICATE1", "CERTIFICATE2"]:
            usecols = ["订单号", "帐号", "金额", "订单状态"]
            dtype = {"订单号": str, "帐号": str, "金额": str, "订单状态": str}

        elif self.subject == "CERTIFICATE4":
            usecols = ["订单时间|订单号", "帐号|姓名", "金额", "状态"]
            dtype = {"订单时间|订单号": str, "帐号|姓名": str, "金额": str, "状态": str}

        else:
            usecols, dtype = None, None
        return usecols, dtype

    def df_data_for_file_filter_condition(self, df_data_for_file):
        df_data_for_file = df_data_for_file[df_data_for_file[self.usecols[3]] == "已处理"]
        return df_data_for_file

    def df_data_cleaning(self, df_data_for_file):
        if self.subject in ["CERTIFICATE1", "CERTIFICATE2"]:
            df_data_for_file["帐号"] = df_data_for_file["帐号"].str.strip()
            df_data_for_file["订单状态"] = df_data_for_file["订单状态"].str.strip()
            df_data_for_file["金额"] = df_data_for_file["金额"].str.extract(r"₫(\d+\.?\d*)k").astype(float)
        elif self.subject == "CERTIFICATE4":
            df_data_for_file["帐号|姓名"] = df_data_for_file["帐号|姓名"].str.strip()
            df_data_for_file["状态"] = df_data_for_file["状态"].str.strip()
            df_data_for_file["订单时间|订单号"] = df_data_for_file["订单时间|订单号"].str.extract(r"订单号：(\S+)")
            df_data_for_file["金额"] = df_data_for_file["金额"].str.extract(r"₫(\d+\.?\d*)k").astype(float)
        return df_data_for_file

    def get_statistical_data(self, df_data_for_file, subset, sum_columns, number_of_members, order_number):
        df_data_for_file.dropna(subset=subset, inplace=True)
        df_data_for_file = self.df_data_cleaning(df_data_for_file)
        df_data_for_file = self.df_data_for_file_filter_condition(df_data_for_file)
        sum_of_amount = round(df_data_for_file[sum_columns].sum(), 4)
        number_of_people = df_data_for_file[number_of_members].nunique()
        number_of_data = df_data_for_file[order_number].count()
        return sum_of_amount, number_of_people, number_of_data


class CertificateForSystem8(CertificateForSystem1):
    def get_read_parameter(self):
        if self.subject == "CERTIFICATE1" or self.subject == "CERTIFICATE2":
            usecols = ["订单号", "会员账号", "金额"]
            dtype = {"订单号": str, "会员账号": str, "金额": str}

        elif self.subject == "CERTIFICATE4":
            usecols = ["订单号", "会员账号", "金额", "平台"]
            dtype = {"订单号": str, "会员账号": str, "金额": str, "平台": str}
        else:
            usecols, dtype = None, None
        return usecols, dtype

    def df_data_for_file_filter_condition(self, df_data_for_file):
        return df_data_for_file


class CertificateForSystem9(CertificateForSystem1):
    def __init__(self, file_path, subject, client_name):
        super().__init__(file_path, subject, client_name)
        self.proportion_column = self.usecols[6]

    def get_read_parameter(self):
        if self.subject == "CERTIFICATE1" or self.subject == "CERTIFICATE2":
            usecols = ["订单号", "会员账号", "订单金额", "订单状态", "会员ID", "成功时间", "币种"]
            dtype = {"订单号": str, "会员账号": str, "订单金额": str, "订单状态": str}

        elif self.subject == "CERTIFICATE4":
            usecols = ["订单号", "会员账号", "金额", "订单状态", "会员ID", "申请时间", "币种"]
            dtype = {"订单号": str, "会员账号": str, "金额": str, "订单状态": str}
        else:
            usecols, dtype = None, None
        return usecols, dtype

    def get_statistical_data(self, df_data_for_file, subset, sum_columns, number_of_members, order_number):
        df_data_for_file.dropna(subset=subset, inplace=True)
        df_data_for_file[self.usecols[2]] = df_data_for_file[self.usecols[2]].astype(float)
        df_data_for_file[self.proportion_column] = df_data_for_file[self.proportion_column].str.extract(
            r'(\d+)').astype(float)
        df_data_for_file[sum_columns] = df_data_for_file[sum_columns] * df_data_for_file[self.proportion_column]
        df_data_for_file = self.df_data_for_file_filter_condition(df_data_for_file)
        sum_of_amount = round(df_data_for_file[sum_columns].sum(), 4)
        number_of_people = df_data_for_file[number_of_members].nunique()
        number_of_data = df_data_for_file[order_number].count()
        return sum_of_amount, number_of_people, number_of_data


class CertificateForSystem10(CertificateForSystem1):
    def get_read_parameter(self):
        if self.subject == "CERTIFICATE1" or self.subject == "CERTIFICATE2":
            usecols = ["充值订单号", "会员账号", "金额", "状态"]
            dtype = {usecols[0]: str, usecols[1]: str, usecols[2]: str, usecols[3]: str}

        elif self.subject == "CERTIFICATE4":
            usecols = ["提现流水号", "申请人登录账号", "金额", "状态"]
            dtype = {usecols[0]: str, usecols[1]: str, usecols[2]: str, usecols[3]: str}
        else:
            usecols, dtype = None, None
        return usecols, dtype

    def df_data_for_file_filter_condition(self, df_data_for_file):
        if self.subject == "CERTIFICATE1" or self.subject == "CERTIFICATE2":
            df_data_for_file = df_data_for_file[df_data_for_file[self.usecols[3]] == "已成功"]
        elif self.subject == "CERTIFICATE4":
            df_data_for_file = df_data_for_file[df_data_for_file[self.usecols[3]] == "已完成"]
        return df_data_for_file


class CertificateForSystem11(CertificateForSystem1):
    def get_read_parameter(self):
        if self.subject == "CERTIFICATE1" or self.subject == "CERTIFICATE2":
            usecols = ["订单号", "会员账号", "金额", "订单状态"]
            dtype = {usecols[0]: str, usecols[1]: str, usecols[2]: str, usecols[3]: str}

        elif self.subject == "CERTIFICATE4":
            usecols = ["订单号", "会员账号", "金额", "订单状态"]
            dtype = {usecols[0]: str, usecols[1]: str, usecols[2]: str, usecols[3]: str}
        else:
            usecols, dtype = None, None
        return usecols, dtype

    def df_data_for_file_filter_condition(self, df_data_for_file):
        if self.subject == "CERTIFICATE1" or self.subject == "CERTIFICATE2":
            df_data_for_file = df_data_for_file[df_data_for_file[self.usecols[3]] == "成功"]
        elif self.subject == "CERTIFICATE4":
            df_data_for_file = df_data_for_file[df_data_for_file[self.usecols[3]].str.contains("成功")]
        return df_data_for_file


class CertificateForSystem12(CertificateForSystem1):
    def get_read_parameter(self):
        if self.subject == "CERTIFICATE1" or self.subject == "CERTIFICATE2":
            usecols = ["订单号", "会员id", "金额", "订单状态"]
            dtype = {usecols[0]: str, usecols[1]: str, usecols[2]: str, usecols[3]: str}

        elif self.subject == "CERTIFICATE4":
            usecols = ["订单号", "会员id", "金额", "订单状态"]
            dtype = {usecols[0]: str, usecols[1]: str, usecols[2]: str, usecols[3]: str}
        else:
            usecols, dtype = None, None
        return usecols, dtype

    def df_data_for_file_filter_condition(self, df_data_for_file):
        if self.subject == "CERTIFICATE1" or self.subject == "CERTIFICATE2":
            df_data_for_file = df_data_for_file[(df_data_for_file[self.usecols[3]] == "已支付") |
                                                (df_data_for_file[self.usecols[3]] == "手動完成")]
        elif self.subject == "CERTIFICATE4":
            df_data_for_file = df_data_for_file[(df_data_for_file[self.usecols[3]] == "已完成")]
        return df_data_for_file


class CertificateForSystem13(CertificateForSystem1):
    def get_read_parameter(self):
        if self.subject == "CERTIFICATE1" or self.subject == "CERTIFICATE2":
            usecols = ["用户名", "用户ID", "金额", "订单状态"]
            dtype = {usecols[0]: str, usecols[1]: str, usecols[2]: str, usecols[3]: str}

        elif self.subject == "CERTIFICATE4":
            usecols = ["用户名", "用户ID", "金额", "订单状态"]
            dtype = {usecols[0]: str, usecols[1]: str, usecols[2]: str, usecols[3]: str}
        else:
            usecols, dtype = None, None
        return usecols, dtype

    def df_data_for_file_filter_condition(self, df_data_for_file):
        if self.subject == "CERTIFICATE1" or self.subject == "CERTIFICATE2":
            df_data_for_file = df_data_for_file[(df_data_for_file[self.usecols[3]] == "成功")]
        elif self.subject == "CERTIFICATE4":
            df_data_for_file = df_data_for_file[(df_data_for_file[self.usecols[3]] == "成功") |
                                                (df_data_for_file[self.usecols[3]] == "强制成功")]
        return df_data_for_file


class CertificateForSystem14(CertificateForSystem1):
    def get_read_parameter(self):
        if self.subject == "CERTIFICATE1" or self.subject == "CERTIFICATE2":
            usecols = ["订单号", "会员账号", "金额", "订单状态"]
            dtype = {usecols[0]: str, usecols[1]: str, usecols[2]: str, usecols[3]: str}

        elif self.subject == "CERTIFICATE4":
            usecols = ["订单号", "会员账号", "预计金额", "订单状态"]
            dtype = {usecols[0]: str, usecols[1]: str, usecols[2]: str, usecols[3]: str}
        else:
            usecols, dtype = None, None
        return usecols, dtype

    def df_data_for_file_filter_condition(self, df_data_for_file):
        if self.subject == "CERTIFICATE1" or self.subject == "CERTIFICATE2":
            df_data_for_file = df_data_for_file[(df_data_for_file[self.usecols[3]] == "成功")]
        elif self.subject == "CERTIFICATE4":
            df_data_for_file = df_data_for_file[(df_data_for_file[self.usecols[3]] == "兑换成功")]
        return df_data_for_file


class Button:
    def __init__(self, location):
        self.location = location

    def get_summery_table(self, inner_date):
        if self.location == "certificate":
            summery_table = SummeryTable(inner_date)
        else:
            summery_table = SummeryTableByDeskTop(inner_date)
        return summery_table

    @staticmethod
    def disable_button():  # 按下取得总表后禁止按钮
        get_summary_table_button.config(state="disabled")
        update_summary_table_button.config(state="disabled")  # 禁用其他按钮
        get_summary_table_button_by_desktop.config(state="disabled")
        update_summary_table_button_by_desktop.config(state="disabled")  # 禁用其他按钮
        update_fail.config(state="disabled")
        update_fail_by_desktop.config(state="disabled")

    @staticmethod
    def enable_button():  # 启用按钮
        get_summary_table_button.config(state="normal")
        update_summary_table_button.config(state="normal")  # 禁用其他按钮
        get_summary_table_button_by_desktop.config(state="normal")
        update_summary_table_button_by_desktop.config(state="normal")  # 禁用其他按钮
        update_fail.config(state="normal")
        update_fail_by_desktop.config(state="normal")

    @staticmethod
    def convert_seconds(seconds):
        hours = seconds // 3600  # 得到小時數
        seconds %= 3600  # 更新剩餘的秒數
        minutes = seconds // 60  # 得到分鐘數
        seconds %= 60  # 更新剩餘的秒數
        return hours, minutes, seconds

    def check_input(self):
        star_time = time.time()
        self.disable_button()
        inner_date = date_entry.get()
        if inner_date == "":
            tk.messagebox.showerror("Error", "請填入日期")
        else:
            try:
                summery_table = self.get_summery_table(inner_date)
                summery_table.get_summary_table()
                style = SummaryStyle(summery_table.file_path)
                style.apply_style()
                end_time = time.time()
                total_time = end_time - star_time
                hours, minutes, seconds = self.convert_seconds(total_time)
                mbox.showinfo("成功", f"完成！總花費{hours}小時{minutes}分{seconds}秒")
            except PermissionError as msg:
                print(msg)
                end_time = time.time()
                total_time = end_time - star_time
                hours, minutes, seconds = self.convert_seconds(total_time)
                tk.messagebox.showerror("失敗", f"檢查總表是否開著，浪費了{hours}小時{minutes}分{seconds}秒")
        self.enable_button()

    def update_summary_table(self):
        star_time = time.time()
        self.disable_button()
        inner_date = date_entry.get()
        if inner_date == "":
            tk.messagebox.showerror("Error", "請填入日期")
        else:
            try:
                summery_table = self.get_summery_table(inner_date)
                summery_table.update_summary_table()
                style = SummaryStyle(summery_table.file_path)
                style.apply_style()
                end_time = time.time()
                total_time = end_time - star_time
                hours, minutes, seconds = self.convert_seconds(total_time)
                mbox.showinfo("成功", f"完成！總花費{hours}小時{minutes}分{seconds}秒")
            except PermissionError as msg:
                print(msg)
                tk.messagebox.showerror("失敗", "檢查總表是否開著")
        self.enable_button()

    def update_fail(self):
        star_time = time.time()
        self.disable_button()
        inner_date = date_entry.get()
        if inner_date == "":
            tk.messagebox.showerror("Error", "請填入日期")
        else:
            try:
                summery_table = self.get_summery_table(inner_date)
                summery_table.update_fail()
                style = SummaryStyle(summery_table.file_path)
                style.apply_style()
                end_time = time.time()
                total_time = end_time - star_time
                hours, minutes, seconds = self.convert_seconds(total_time)
                mbox.showinfo("成功", f"完成！總花費{hours}小時{minutes}分{seconds}秒")
            except PermissionError as msg:
                print(msg)
                tk.messagebox.showerror("Error", "檢查日期、資料夾中是否有其他檔案")
        self.enable_button()


class RedirectText(object):
    def __init__(self, text_ctrl):
        self.output = text_ctrl

    def write(self, string):
        self.output.insert(tk.END, string)
        self.output.see(tk.END)

    def flush(self):
        pass


if __name__ == "__main__":
    multiprocessing.freeze_support()
    window = tk.Tk()
    window.title("充值提線-收支")
    window.geometry('600x400')

    account_day_range = tk.StringVar(window)
    account_day_range.set("")
    date_entry = tk.Entry(window, textvariable=account_day_range)
    date_entry.grid(row=0, column=0, columnspan=4)
    # -------
    get_summary_table_button = tk.Button(window, text="取得總表",
                                         command=lambda: threading.Thread(
                                             target=Button("certificate").check_input).start())
    get_summary_table_button.grid(row=1, column=0)
    # -------
    update_summary_table_button = tk.Button(window, text="更新",
                                            command=lambda: threading.Thread(
                                                target=Button("certificate").update_summary_table).start())
    update_summary_table_button.grid(row=1, column=1)
    # -------
    update_fail = tk.Button(window, text="更新失敗",
                            command=lambda: threading.Thread(target=Button("certificate").update_fail).start())
    update_fail.grid(row=1, column=2)
    # -------
    get_summary_table_button_by_desktop = tk.Button(window, text="取得總表(桌面)",
                                                    command=lambda: threading.Thread(
                                                        target=Button("desktop").check_input).start())
    get_summary_table_button_by_desktop.grid(row=2, column=0)
    # -------
    update_summary_table_button_by_desktop = tk.Button(window, text="更新(桌面)",
                                                       command=lambda: threading.Thread(
                                                           target=Button("desktop").update_summary_table).start())
    update_summary_table_button_by_desktop.grid(row=2, column=1)
    # -------
    update_fail_by_desktop = tk.Button(window, text="更新失敗(桌面)",
                                       command=lambda: threading.Thread(target=Button("desktop").update_fail).start())
    update_fail_by_desktop.grid(row=2, column=2)


    text_area = tk.Text(window)
    text_area.grid(row=3, column=0, columnspan=4)

    window.mainloop()
