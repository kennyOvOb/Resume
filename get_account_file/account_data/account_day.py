from typing import Tuple
from datetime import date, timedelta

RIGHT_NOW = date.today()  # 抓今天時間


def get_account_day() -> Tuple[date, int, int, int, str, str]:  # 取得當前日期的帳務日各種資訊
    account_day = (RIGHT_NOW - timedelta(days=2))  # 目前帳務日 ，有個問題是01 02 03，日期格式
    account_year = account_day.year
    account_month = account_day.month
    account_day_by_day = account_day.day  # 抓幾天
    account_day_str_md = account_day.strftime("%m%d")  # 改成月月日日格式
    account_day_str_ymd = account_day.strftime("%y%m%d")  # 改成年年月月日日格式
    return account_day, account_year, account_month, account_day_by_day, account_day_str_md, account_day_str_ymd

