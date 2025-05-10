from typing import Tuple
from datetime import date, timedelta

RIGHT_NOW = date.today()  # 抓今天時間


def get_account_day() -> Tuple[date, int, int, int, str, str]:
    account_day = (RIGHT_NOW - timedelta(days=2))
    account_year = account_day.year
    account_month = account_day.month
    account_day_by_day = account_day.day  # 抓幾天
    account_day_str_md = account_day.strftime("%m%d")  # 改成月月日日格式
    account_day_str_ymd = account_day.strftime("%y%m%d")  # 改成年年月月日日格式
    return account_day, account_year, account_month, account_day_by_day, account_day_str_md, account_day_str_ymd

