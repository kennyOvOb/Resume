import sys
# import logging
import threading
import asyncio
import time
import tkinter as tk
import pandas as pd
from tkinter import messagebox
from datetime import datetime, timedelta
from pathlib import Path
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, CallbackContext, filters




class QuestionBot:
    def __init__(self):
        self.token_api: str = "your_bot_token"  # 機器人token
        self.ga_data_path = Path(r"Cloud_Drive_path")
        self.account_common_path = self.ga_data_path / "06-共用資料" / "06-02-帳務共用"
        self.account_common_tool_path = self.account_common_path / "06-02-05-共用工具"
        self.bot_file_path = self.account_common_tool_path / "08-提問機器人"
        self.database_path: Path = self.account_common_tool_path / "數據庫-2022.08啟用.xlsx"  # 客戶Telegram群組id對應表路徑
        self.group_message_send_file_path: Path = self.bot_file_path / "群發訊息.xlsx"
        self.group_message_sub_file_path: Path = self.bot_file_path / "群發檔案"
        self.datatime_now: datetime = datetime.utcnow() # noqa
        # 取得機器人創建的時間，也就是bot執行時間，後續為了只執行機器人啟動時的指令，因telegram紀錄時間為uct+0故使用uct now
        self.today_account_str: str = (self.datatime_now - timedelta(2)).strftime(
            '%y%m%d')  # 帳務日為今天日期減二，問題格式為240101-客戶名-提問.xlsx，因此需要字串
        self.fail_doc: list = list()  # 失敗檔案的list
        self.fail_message_site: list = list()  # 失敗訊息的list
        self.group_id_dict: dict = self.get_group_id_dict()  # 使用對應表取的字典 {客戶名:群組ID}
        self.available_users: list = self.get_available_users_list()  # 使用對應表取得允許的使用者，公司所使用帳號
        self.inner_question_amount = None  # 問題數量，預設為空，由使用者輸入，用來確認問題數量
        self.count: int = 0  # 問題計數器，後續跌代使用，到達問題數量觸發end傳送訊息
        self.group_message_columns: list = ["客戶", "群發訊息", "檔名"]
        self.id_message_count = dict()
        self.sub_msg = ""
        self.title_mode = False
        self.merge_mode = False
        self.file_mode = False
        self.group_message_df = None
        self.mode_settings = {
            "T": self.set_title_mode,
            "M": self.set_merge_mode,
            "F": self.set_file_mode,
        }
        self.application = ApplicationBuilder().token(self.token_api).build()  # 機器人創建
        self.set_command_handler()  # 設定機器人指令

    def get_group_message_info(self) -> (int, list, list):
        """
        :return:返回群組:需發送訊息的字典
        """
        subset = self.group_message_columns[:2]
        try:
            # noinspection PyTypeChecker
            df = pd.read_excel(self.group_message_send_file_path, sheet_name="群發訊息",
                               usecols=self.group_message_columns,
                               dtype={"客戶": str, "群發訊息": str, "檔名": str}).dropna(subset=subset)
            df["檔名"].fillna("", inplace=True)
            self.group_message_df = df
        except Exception:
            raise
        message_count: int = len(df)
        client_name_list = df["客戶"].tolist()
        message_list = df["群發訊息"].tolist()
        return message_count, client_name_list, message_list

    def get_available_users_list(self) -> list:
        """
        取得允許的使用者，公司內部telegram帳號
        :return: 回傳使用者ID的list，因為後許僅用in確認是否在列表中
        """
        # noinspection PyTypeChecker
        df: pd.DataFrame = pd.read_excel(self.database_path, sheet_name="群組ID", usecols=["使用者", "使用者ID"],
                                         dtype={"使用者": str, "使用者ID": str}).dropna(subset="使用者ID")

        return df["使用者ID"].to_list()

    def get_group_id_dict(self) -> dict:
        """
        取得客戶對應群組ID的字典
        將客戶名改為大寫
        {客戶名:群組ID}
        :return:df["群組ID"].to_dict() 回傳dict
        """
        # noinspection PyTypeChecker
        df: pd.DataFrame = pd.read_excel(self.database_path, sheet_name="群組ID", usecols=["客戶", "群組ID"],
                                         dtype={"客戶": str, "群組ID": str}).dropna(
            subset="群組ID")
        df["客戶"] = df["客戶"].str.upper()
        df.set_index("客戶", inplace=True)
        return df["群組ID"].to_dict()

    def time_check(self, message_time) -> bool:
        """
        將massage轉為uct+0
        :param message_time: 訊息傳送時間
        :return:self.datatime_now > message_time 回傳創建時間是否大於訊息時間
        """
        message_time = message_time.replace(tzinfo=None)
        # print(self.datatime_now)
        # print(message_time)
        # print(self.datatime_now > message_time)
        return self.datatime_now > message_time

    def user_check(self, chat_id) -> bool:
        """
        :param chat_id: 傳訊息的人ID
        :return:chat_id not in self.available_users 回傳是否沒在允許的使用著list中
        """
        # print(chat_id)
        # print(self.available_users)
        # print(str(chat_id) not in self.available_users)
        return str(chat_id) not in self.available_users

    @staticmethod
    def type_check(chat_type) -> bool:
        """
        :param chat_type:聊天的型態 會是 ‘private’, ‘group’, ‘supergroup’ or ‘channel’這四種
        :return:chat_type != "private" 判別是否不是私人訊息
        """
        # print(chat_type != "private")
        return chat_type != "private"

    async def question_count_check(self, update: Update, context: CallbackContext):
        """
        chat_id:傳訊息的人
        message_time:訊息時間
        chat_type:聊天的類型
        檢查時間、檢查時間使用者、檢查類型，只要有任一成立，直接回傳中斷
        telegram回傳訊息，告知設定問題的數量
        :param update:更新器object
        :param context:回傳訊息object
        :return:不回傳
        """
        chat_id = update.effective_chat.id  # 傳訊息的人
        message_time = update.effective_message.date
        chat_type = update.effective_chat.type
        if self.time_check(message_time) or self.user_check(chat_id) or self.type_check(chat_type):
            return
        await context.bot.send_message(chat_id=chat_id,
                                       text=f"提問數量為: {self.inner_question_amount}")

    async def get_id(self, update: Update, context: CallbackContext):
        """
        chat_id:傳訊息的人
        message_time:訊息時間
        檢查時間、若成立直接回傳中斷，只檢查時間是因為需要在群組使用，獲取群組ID
        telegram回傳訊息，告知ID資訊
        :param update:更新器object
        :param context:回傳訊息object
        :return: 不回傳
        """
        chat_id = update.effective_chat.id  # 傳訊息的人
        message_time = update.effective_message.date
        if self.time_check(message_time):
            return
        await context.bot.send_message(chat_id=chat_id, text=f"ID資訊為: {chat_id}")

    async def question_count_reset(self, update: Update, context: CallbackContext):
        """
        chat_id:傳訊息的人
        message_time:訊息時間
        chat_type:聊天的類型
        檢查時間、檢查時間使用者、檢查類型，只要有任一成立，直接回傳中斷
        msg = 取得此次訊息內容
        try
        指令為:/指令 空白 訊息，為避免有人打空白或全形，因此替換
        分列空白取得，指令後方字串
        設定問題數量為指定數量
        並回傳訊息告知，已重設
        若上述發生錯誤，則告知重設失敗
        :param update:更新器object
        :param context:回傳訊息object
        :return:不回傳
        """
        chat_id = update.effective_chat.id  # 傳訊息的人
        message_time = update.effective_message.date
        chat_type = update.effective_chat.type
        if self.time_check(message_time) or self.user_check(chat_id) or self.type_check(chat_type):
            return
        msg = update.effective_message
        try:
            normalized_text = msg.text.replace('　', ' ')  # 怕有人打成全形
            command, number_str = normalized_text.split(' ')  # split the command and the number
            self.inner_question_amount = int(number_str)
            self.group_id_dict: dict = self.get_group_id_dict()  # 使用對應表取的字典 {客戶名:群組ID}
            self.available_users: list = self.get_available_users_list()  # 使用對應表取得允許的使用者，公司所使用帳號
            await context.bot.send_message(chat_id=chat_id,
                                           text=f"已重設問題數量為: {self.inner_question_amount}")
        except Exception:  # noqa
            await context.bot.send_message(chat_id=chat_id,
                                           text=f"重設失敗")

    async def process_end(self, fail_var_str, update: Update, context: CallbackContext):
        """
        沒檢查時間、使用者、類型是因為此end只會接續在轉傳後，問題數量滿足設定的數量後才會執行
        chat_id:傳訊息的人
        將所記錄的失敗文件，以\n換行串聯，使得傳訊息整齊排列
        計算失敗數量
        計算成功數量
        若有失敗檔案則傳送計數跟失敗檔案
        若無則告直沒有失敗檔案
        最後再重設失敗檔案列表跟失敗計數器
        :param fail_var_str: 失敗列表變數名字串
        :param update:更新器object
        :param context:回傳訊息object
        :return: 不回傳
        """
        if fail_var_str == "fail_doc":
            message_type = "檔案"
        elif fail_var_str == "fail_message_site":
            message_type = "訊息"
        else:
            print("錯誤")
            return

        fail_list = getattr(self, fail_var_str)  # 用傳入的變數字串取得該變數的值
        chat_id = update.effective_chat.id  # 傳訊息的人
        fail_text: str = "\n".join(fail_list)
        fail_count: int = len(fail_list)
        success_count: int = self.count - fail_count

        if fail_text != "":
            await context.bot.send_message(chat_id=chat_id, text=f"總計傳送{self.count}{message_type}"
                                                                 f"\n成功:{success_count}個，失敗{fail_count}個"
                                                                 f"\n失敗{message_type}如下:\n{fail_text}")
        else:
            await context.bot.send_message(chat_id=chat_id,
                                           text=f"總計傳送{self.count}{message_type}\n無失敗{message_type}")
        setattr(self, fail_var_str, list())  # 用傳入的變數字串重設變數的值
        self.count: int = 0
        self.id_message_count = dict()

    def clean_group_message_file(self):
        """
        將群發訊息檔案重新保存空的一份，做到清空
        :return:None
        """
        empty_df = pd.DataFrame(columns=self.group_message_columns)
        with pd.ExcelWriter(self.group_message_send_file_path) as writer:
            empty_df.to_excel(writer, "群發訊息", index=False)

    def send_times_check(self, chat_id):
        chat_id_str = str(chat_id)
        chat_id_count = self.id_message_count.get(chat_id_str, 0)
        if chat_id_count % 20 == 0 and chat_id_count != 0:
            time.sleep(60)
        if self.count % 30 == 0:
            time.sleep(1)

    def get_message_merge_dict(self):
        message_count, client_name_list, message_list = self.get_group_message_info()
        group_message_dict: dict = dict()
        for count in range(message_count):
            client_name = client_name_list[count]
            message = message_list[count]
            if client_name not in group_message_dict:
                group_message_dict[client_name] = message
            else:
                group_message_dict[client_name] = group_message_dict[client_name] + "\n" + message

        return group_message_dict

    def massage_add_client_name(self, client_name, message):
        if self.sub_msg != "":
            full_msg = client_name + "\n" + self.sub_msg + "\n" + message
        else:
            full_msg = client_name + "\n" + message
        return full_msg

    async def send_separate_message(self):
        message_count, client_name_list, message_list = self.get_group_message_info()
        self.inner_question_amount = message_count
        for count in range(message_count):
            client_name = client_name_list[count]
            message = message_list[count]
            await self.send_message(client_name, message)

    async def send_merge_message(self):
        group_message_dict = self.get_message_merge_dict()
        self.inner_question_amount = len(group_message_dict)
        for client_name, message in group_message_dict.items():
            await self.send_message(client_name, message)

    def update_id_message_count(self, chat_id):
        chat_id_str = str(chat_id)
        if chat_id_str in self.id_message_count:
            self.id_message_count[chat_id_str] += 1
        else:
            self.id_message_count[chat_id_str] = 1

    def get_sub_file_path(self, client_name, message):
        """
        黨訊息要附加檔案時，依照客戶與訊息去找到群發訊息檔案中紀錄的對應檔案路徑
        :param client_name: 客戶名稱
        :param message: 訊息
        :return:
        """
        mask = (self.group_message_df["客戶"] == client_name) & (self.group_message_df["群發訊息"] == message)
        file_name_list = self.group_message_df.loc[mask, "檔名"].values
        if len(file_name_list) == 1:
            file_name = file_name_list[0]
            if file_name == "":
                return None
            sub_file_path = self.group_message_sub_file_path / file_name
            if sub_file_path.is_file():
                return sub_file_path
            raise ValueError(client_name + " " + file_name + "，檔名有誤")
        elif len(file_name_list) > 1:
            raise ValueError(message + "\n" + "訊息重複")
        return None

    async def send_message(self, client_name, message):
        group_id = self.get_group_id(client_name)

        if self.title_mode:
            full_message = self.massage_add_client_name(client_name, message)
        else:
            full_message = message

        try:
            if self.file_mode:
                sub_file_path = self.get_sub_file_path(client_name, message)
                if sub_file_path is not None:
                    await self.application.bot.send_document(chat_id=group_id, caption=full_message,
                                                             document=sub_file_path)
                else:
                    await self.application.bot.send_message(chat_id=group_id, text=full_message)
            else:
                await self.application.bot.send_message(chat_id=group_id, text=full_message)

            self.update_id_message_count(group_id)
        except Exception as msg:  # noqa
            self.fail_message_site.append(client_name + "\n" + message + "\n")
            print(client_name)
            print(msg)
        finally:
            self.count += 1
            self.send_times_check(group_id)

    def reset_mode(self):
        self.sub_msg = ""
        self.title_mode = False
        self.merge_mode = False
        self.file_mode = False

    def set_title_mode(self, text_split):
        self.title_mode = True
        self.sub_msg = ' '.join(text_split[1:])

    def set_merge_mode(self, _):
        self.merge_mode = True

    def set_file_mode(self, _):
        self.file_mode = True

    async def gms_parameter_parse(self, update: Update, context: CallbackContext):
        """
        解析群發訊息的參數
        :param update:
        :param context:
        :return:
        """
        msg = update.effective_message
        chat_id = update.effective_chat.id
        error = False
        try:
            normalized_text = msg.text.upper().replace('　', ' ')  # 變成大寫，怕有人打成全形，替換後刪除空格
            normalized_text_split = normalized_text.split("-")
            text_count = len(normalized_text_split)
            if text_count == 1:
                if normalized_text.replace(' ', "") == "/GMS":
                    return error

            for text in normalized_text_split:
                text_split = text.split(' ')
                first_text = text_split[0]
                if first_text == "/GMS":
                    continue

                setting = self.mode_settings[first_text]

                if setting is not None:
                    setting(text_split) # noqa
                    continue

                raise ValueError("命令錯誤")

            if self.merge_mode is True and self.file_mode is True:
                error = True
                raise ValueError("命令錯誤")

        except Exception:  # noqa
            await context.bot.send_message(chat_id=chat_id,
                                           text=f"命令錯誤")
            self.reset_mode()
            error = True
        finally:
            return error

    async def group_message_send(self, update: Update, context: CallbackContext):
        """
        依照群發訊息excel檔案資料，群發訊息
        :param update:更新器object
        :param context:回傳訊息object
        :return: None
        """
        self.group_id_dict: dict = self.get_group_id_dict()  # 使用對應表取的字典 {客戶名:群組ID}
        self.available_users: list = self.get_available_users_list()  # 使用對應表取得允許的使用者，公司所使用帳號
        user_id = update.effective_chat.id
        message_time = update.effective_message.date
        chat_type = update.effective_chat.type

        if self.time_check(message_time) or self.user_check(user_id) or self.type_check(chat_type):
            return
        error = await self.gms_parameter_parse(update, context)
        if not error:
            try:
                if self.merge_mode:
                    await self.send_merge_message()
                else:
                    await self.send_separate_message()

                await self.process_end("fail_message_site", update, context)
            except Exception as msg:  # noqa
                print(msg)
                await context.bot.send_message(chat_id=user_id, text="群發訊息檔案讀取失敗")
            try:
                self.clean_group_message_file()

            except Exception:  # noqa
                await context.bot.send_message(chat_id=user_id, text="群發訊息檔案保存失敗，請確認是否開啟以及留意清空")
            finally:
                self.reset_mode()

    async def help_info(self, update: Update, context: CallbackContext):
        chat_id = update.effective_chat.id  # 傳訊息的人
        message_time = update.effective_message.date
        chat_type = update.effective_chat.type
        if self.time_check(message_time) or self.user_check(chat_id) or self.type_check(chat_type):
            return

        await context.bot.send_message(chat_id=chat_id, text="目前有以下指令\n"
                                                             "help，取得指令資訊\n"
                                                             "check，取得目前設定的問題數量\n"
                                                             "ID，取得提交命令的群組或使用者ID\n"
                                                             "reset，重設問題數量以及重新取得資料庫群組ID資料\n"
                                                             "gms，依照群發訊息檔案中資料群發消息，有以下兩個指令參數\n"
                                                             "-t，會在傳送的訊息首行新增客戶名稱\n"
                                                             "參數後方空白可銜接額外備註，例如-t 範例說明\n"
                                                             "-m，若群發訊息檔案中是同一客戶有好幾行訊息，會合併為一條一次傳送\n"
                                                             "若有參數請以空白間隔")

    def get_group_id(self, client_name) -> str or None:
        """
        使用客戶名取得self.group_id_dict中對應的群組ID
        :param client_name: 客戶名稱
        :return: 群組ID，沒找到回傳None
        """
        return self.group_id_dict.get(client_name.upper(), None)

    async def document_handler(self, update: Update, context: CallbackContext):
        """
        chat_id:傳訊息的人
        message_time:訊息時間
        chat_type:聊天的類型
        檢查時間、檢查時間使用者、檢查類型，只要有任一成立，直接回傳中斷
        doc:取的文件內容
        doc_split，依照"-"分列，問題格式固定為日期-客戶名稱-提問.xlsx
        date_str:取得日期字串
        取得客戶名，之所以不是用doc_split[1]，是因為客戶名也有"-"，因此使用"-"join，範圍是[1:-1]也就是去掉日期跟提問.xlsx
        如果日期等於帳務日跟有找到群組ID，則轉傳此問題到該群組，計數+1
        如出錯告知錯誤原因，列出問題檔案，計數器+1
        如果日期不對或找不到ID，列出問題檔案，計數器+1
        最後檢查計數是否等於設定問題數量，滿足執行self.end
        :param update:更新器object
        :param context:回傳訊息object
        :return: 不回傳
        """
        chat_id = update.effective_chat.id  # 傳訊息的人
        message_time = update.effective_message.date
        chat_type = update.effective_chat.type
        if self.time_check(message_time) or self.user_check(chat_id) or self.type_check(chat_type):
            return

        doc = update.effective_message.document  # 提取文件
        doc_split: list = doc.file_name.split('-')
        date_str = doc_split[0]
        client_name: str = "-".join(doc_split[1:-1])  # 分割文件名稱以獲取客戶名稱
        group_id = self.get_group_id(client_name)
        if self.today_account_str == date_str and group_id:
            try:
                await context.bot.forward_message(chat_id=group_id,
                                                  from_chat_id=chat_id,
                                                  message_id=update.effective_message.message_id)
                self.update_id_message_count(group_id)
            except Exception as msg:
                print(msg)
                self.fail_doc.append(doc.file_name)

                print(doc.file_name)
            finally:
                self.count += 1
                self.send_times_check(group_id)

        else:
            self.fail_doc.append(doc.file_name)
            print(doc.file_name)
            self.count += 1

        if self.count == self.inner_question_amount:
            await self.process_end("fail_doc", update, context)

    def set_command_handler(self):
        """
        設定指令
        :return:不回傳
        """
        start_handler = CommandHandler('check', self.question_count_check)
        id_handler = CommandHandler('ID', self.get_id)
        reset_handler = CommandHandler('reset', self.question_count_reset)
        group_send_handler = CommandHandler('gms', self.group_message_send)
        help_info_handler = CommandHandler('help', self.help_info)
        self.application.add_handler(start_handler)
        self.application.add_handler(id_handler)
        self.application.add_handler(reset_handler)
        self.application.add_handler(group_send_handler)
        self.application.add_handler(help_info_handler)
        # 只有傳送檔案是xlsx才會觸發此指令
        self.application.add_handler(MessageHandler(filters.Document.FileExtension("xlsx"), self.document_handler))

    def main(self, inner_question_amount):
        """
        重設問題計數器數量為0
        設定循環事件
        :param inner_question_amount:要設定的問題數量
        :return: 不回傳
        """
        self.count: int = 0
        self.inner_question_amount = inner_question_amount
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        try:
            self.application.run_polling()
        finally:
            loop.close()


class RedirectText(object):
    def __init__(self, text_ctrl):
        self.output = text_ctrl

    def write(self, string):
        self.output.insert(tk.END, string)
        self.output.see(tk.END)

    def flush(self):
        pass


class Button:
    def __init__(self):
        pass

    @staticmethod
    def start():

        try:
            inner_question_amount: int = int(question_amount.get())
            if inner_question_amount > 0:
                print("star")
                bot_online_button.config(state="disabled")
                thread = threading.Thread(target=lambda: bot.main(inner_question_amount))
                thread.daemon = True  # Set the thread as a daemon
                thread.start()  # 已经在前边创建了thread，这里只需要调用start方法
            else:
                tk.messagebox.showerror("Error", "請輸入大於0的數字")
        except ValueError:
            tk.messagebox.showerror("Error", "請輸入數字")


if __name__ == '__main__':
    window = tk.Tk()
    bot = QuestionBot()
    window.title("問題機器人")
    window.geometry('600x350')
    text_area1 = tk.Text(window)
    text_area1.grid(row=1, column=0, columnspan=4)

    # 重新抓取控制台输出
    old_stdout = sys.stdout
    old_stderr = sys.stderr
    sys.stdout = RedirectText(text_area1)
    sys.stderr = RedirectText(text_area1)

    question_amount = tk.StringVar(window)
    question_amount.set("")
    question_amount_entry = tk.Entry(window, textvariable=question_amount)
    question_amount_entry.grid(row=0, column=0, columnspan=1)

    bot_online_button = tk.Button(window, text="啟動機器人", command=Button.start)
    bot_online_button.grid(row=0, column=2)

    window.mainloop()
    bot.application.shutdown()
    sys.stdout = old_stdout
    sys.stderr = old_stderr
