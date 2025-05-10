import pandas as pd
from pathlib import Path


class Database:
    def __init__(self):
        self.sheet_name = "數據庫"
        self.database_path: Path = Path(
            r"Could_drive_path")
        self.site_name_number_dict = self.get_site_name_number_dict()
        self.site_name_number_upper_dict = {str(name).upper(): number for name, number in self.site_name_number_dict.items()}

    def get_site_name_number_dict(self):
        """
        盤口名稱為大寫
        :return:
        """
        df = pd.read_excel(self.database_path, sheet_name=self.sheet_name, usecols=["客戶名称", "编号"], header=1) # noqa
        df["客戶名称"] = df["客戶名称"].str.upper()
        df.set_index("客戶名称", inplace=True)
        site_name_number_dict = df.to_dict()['编号']
        return site_name_number_dict

