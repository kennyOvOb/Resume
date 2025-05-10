from datetime import datetime
from pathlib import Path
import pandas as pd
from pandas import ExcelFile


class Distribution:
    def __init__(self, account_date):
        self.account_datetime = account_date
        self.account_datetime_last_month = datetime(self.account_datetime.year, self.account_datetime.month - 1, 1)
        self.distribution_path: Path = Path(
            r"could_drive_path")
        self.sheet_name = self.account_datetime.strftime("%y-%#m")
        self.sheet_name_last_month = self.account_datetime_last_month.strftime("%y-%#m")
        self.distribution_df = pd.read_excel(self.distribution_path, sheet_name=self.sheet_name)
        self.distribution_last_month_df = pd.read_excel(self.distribution_path, sheet_name=self.sheet_name_last_month)
        self.number_site_name_dict_two_month = self.get_number_site_name_dict_two_month()
        self.distribute_note = self.get_distribute_note()
        self.distribute_for_person_dict = self.get_distribute_for_person_dict()
        self.name_number_dict = self.get_name_number_dict()


    def get_client_info_df(self, df_var):
        # noinspection PyTypeChecker
        distribution_df_one_site = getattr(self, df_var).copy().iloc[:, [1, 2, 4, 6]]
        # noinspection PyTypeChecker
        distribution_df_total_site = getattr(self, df_var).copy().iloc[:, [15, 16, 18, 20]]
        distribution_df_total_site.columns = ["客戶名稱", "编码", "系统", "币别"]
        total_client_info_df = pd.concat([distribution_df_one_site, distribution_df_total_site])
        return total_client_info_df

    def get_number_site_name_dict(self, df_var):
        """
        盤口名稱為大寫
        :param df_var:df變數名稱
        :return:
        """
        df = self.get_client_info_df(df_var)[["客戶名稱", "编码"]]
        df.set_index("编码", inplace=True)
        number_site_name_dict = df.to_dict()['客戶名稱']
        return number_site_name_dict

    def get_number_site_name_dict_two_month(self):
        """
        盤口名稱為大寫
        :return:
        """
        number_site_name_dict1 = self.get_number_site_name_dict("distribution_last_month_df")
        number_site_name_dict2 = self.get_number_site_name_dict("distribution_df")
        number_site_name_dict1.update(number_site_name_dict2)
        return number_site_name_dict1

    def get_distribute_for_person_dict(self):
        def system_rename(df):
            main_system: list = ["system1", "system2", "system3"]
            for system in main_system:
                mask = df["系统"].str.contains(system, na=False)
                df.loc[mask, "系统"] = system
            mask = ~df["系统"].str.contains('|'.join(main_system), na=False)
            df.loc[mask, "系统"] = "其他"
            return df

        distribute_for_person_dict = {}
        suffix_map = {"別組1": "-模板", "別組2": "-模板", "別組3": "-模板", "共同处理": "-彙總", "汇总": "歸檔",
                      "系统": "系统"}
        # 各人員負責清單
        mould_groups_name = ["別組1", "別組2", "別組3"]
        # 彙總清單
        summary_groups_name = ["共同处理"]
        # 歸檔清單
        induction_groups_name = ["汇总"]
        # 系統清單
        system_groups_name = ["系统"]
        total_groups_name = mould_groups_name + summary_groups_name + induction_groups_name + system_groups_name
        try:
            excel_file = ExcelFile(self.distribution_path)
            all_sheets = excel_file.sheet_names
            if self.sheet_name in all_sheets:
                # noinspection PyTypeChecker
                distribution_df = self.distribution_df.copy().iloc[:, 0:12].dropna(subset="客戶名稱")
                # 因為系統名不一致，需重新命名一致
                distribution_df = system_rename(distribution_df)
                for group_name in total_groups_name:
                    group_by_df = distribution_df.groupby(group_name)
                    for name, group in group_by_df:
                        list_title = str(name) + suffix_map.get(group_name)
                        distribute_for_person_dict[list_title] = group["客戶名稱"].tolist()
                distribute_for_person_dict["自定清單"] = []
            else:
                distribute_for_person_dict = {"分配表無此月份分頁": []}
        except KeyError:
            distribute_for_person_dict = {"分配表無此月份分頁": []}
        finally:
            return distribute_for_person_dict

    def get_distribute_note(self):
        # noinspection PyTypeChecker
        distribution_note_df_one_site = self.distribution_df.copy().iloc[:, [1, 13]].dropna(subset= "客戶名稱")
        # noinspection PyTypeChecker
        distribution_note_df_one_site["客戶名稱"] = distribution_note_df_one_site["客戶名稱"].str.upper()
        distribution_note_df_one_site.set_index("客戶名稱", inplace=True)
        distribution_note_df_one_site_dict = distribution_note_df_one_site.to_dict()["備註"]
        return distribution_note_df_one_site_dict

    def get_name_number_dict(self):
        # noinspection PyTypeChecker
        name_upper_number_df_one_site = self.distribution_df.copy().iloc[:, [1, 2]]
        # noinspection PyTypeChecker
        name_upper_number_df_one_site["客戶名稱"] = name_upper_number_df_one_site["客戶名稱"]
        name_upper_number_df_one_site.set_index("客戶名稱", inplace=True)
        name_number_dict = name_upper_number_df_one_site.to_dict()["编码"]
        return name_number_dict

