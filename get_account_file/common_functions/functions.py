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


def get_last_day(str_date_ymd):
    import calendar
    year = int(str_date_ymd[0:2])
    month = int(str_date_ymd[2:4])
    _, last_day = calendar.monthrange(year, month)
    return last_day
