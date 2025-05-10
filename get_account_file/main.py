import tkinter as tk
from account_data.account_day import get_account_day
from getcertificate.get_today_file import TodayFile
from multiprocessing import Process, freeze_support
import threading


class RedirectText(object):
    def __init__(self, text_ctrl):
        self.output = text_ctrl

    def write(self, string):
        self.output.insert(tk.END, string)

    def flush(self):
        pass  # 在此处不需要做任何操作


def update_user_list(*args): # noqa
    global option_menu
    global user_list
    global today_file

    str_date_ymd = date_entry.get()
    if len(str_date_ymd) == 6:
        option_menu.place_forget()
        today_file = TodayFile(str_date_ymd)
        user_list = list(today_file.distribution.distribute_for_person_dict.keys())
        print(f"new user_list: {user_list}")
        choice.set("選擇使用者")

        option_menu = tk.OptionMenu(window, choice, *user_list)
        option_menu.place(x=150, y=0)


def execute_checked_tasks(user, today_file_obj):
    disable_button()
    window.update()
    tasks = []
    task_list = [
        (get_today_file_check, "mould_board"),
        (get_today_account_check, "today_account"),
        (get_summary_check, "summary"),
        (get_question_check, "question"),
        (get_over_board_check, "over_mould_board"),
        (get_over_summary_check, "over_summary")]

    for check, action in task_list:
        if check.get():
            task_process = Process(target=today_file_obj.get_file, args=(user, action))
            task_process.start()
            tasks.append(task_process)

    window.after(1000, check_tasks, tasks)


def check_tasks(tasks):
    all_done = all(not task.is_alive() for task in tasks)
    if not all_done:
        # If not all tasks are done, check again after 100ms.
        window.after(3, check_tasks, tasks)
    else:
        enable_button()


def configure_buttons(state):
    button_list = [
        date_entry, option_menu, get_today_file_check_button, get_today_file_button,
        get_today_account_check_button, get_today_account_button, get_summary_check_button,
        get_summary_button, get_question_check_button, get_question_button, execute_checked_button,
        get_over_board_check_button, get_over_board_button, get_over_summary_check_button,
        get_over_summary_button
    ]

    for button in button_list:
        button.config(state=state)


def disable_button():
    configure_buttons("disabled")


def enable_button():
    configure_buttons("normal")


if __name__ == "__main__":
    freeze_support()
    window = tk.Tk()
    window.title("Account Management")
    window.geometry('280x170')
    window.resizable(False, False)

    account_day_str_ymd = tk.StringVar(window)
    account_day_str_ymd.set(get_account_day()[5])
    account_day_str_ymd.trace("w", update_user_list)
    date_entry = tk.Entry(window, textvariable=account_day_str_ymd)
    date_entry.place(x=0, y=5)

    today_file = TodayFile()
    user_list = list(today_file.distribution.distribute_for_person_dict.keys())
    choice = tk.StringVar(window)
    choice.set("選擇使用者")
    option_menu = tk.OptionMenu(window, choice, *user_list)
    option_menu.place(x=150, y=0)

    get_today_file_check = tk.IntVar(value=0)
    get_today_file_check_button = tk.Checkbutton(window, variable=get_today_file_check)
    get_today_file_check_button.place(x=10, y=30)

    get_today_file_button = tk.Button(window, text="取得今天模板",
                                      command=lambda: today_file.get_file(choice.get(), "mould_board"))
    get_today_file_button.place(x=40, y=30)

    get_today_account_check = tk.IntVar(value=0)
    get_today_account_check_button = tk.Checkbutton(window, variable=get_today_account_check)
    get_today_account_check_button.place(x=130, y=30)

    get_today_account_button = tk.Button(window, text="取得當日帳務",
                                         command=lambda: today_file.get_file(choice.get(), "today_account"))
    get_today_account_button.place(x=160, y=30)


    get_summary_check = tk.IntVar(value=0)
    get_summary_check_button = tk.Checkbutton(window, variable=get_summary_check)
    get_summary_check_button.place(x=10, y=60)

    get_summary_button = tk.Button(window, text="取得今日彙總",
                                   command=lambda: today_file.get_file(choice.get(), "summary"))
    get_summary_button.place(x=40, y=60)

    get_question_check = tk.IntVar(value=0)
    get_question_check_button = tk.Checkbutton(window, variable=get_question_check)
    get_question_check_button.place(x=130, y=60)

    get_question_button = tk.Button(window, text="取得問題回復",
                                    command=lambda: today_file.get_file(choice.get(), "question"))
    get_question_button.place(x=160, y=60)

    execute_checked_button = tk.Button(window, text="執行勾選任務",
                                       command=lambda: threading.Thread(
                                           execute_checked_tasks(choice.get(), today_file), daemon=True))
    execute_checked_button.place(x=90, y=120)


    # ----

    get_over_board_check = tk.IntVar(value=0)
    get_over_board_check_button = tk.Checkbutton(window, variable=get_over_board_check)
    get_over_board_check_button.place(x=10, y=90)

    get_over_board_button = tk.Button(window, text="取得關帳模板",
                                      command=lambda: today_file.get_file(choice.get(), "over_mould_board"))
    get_over_board_button.place(x=40, y=90)

    # ---
    get_over_summary_check = tk.IntVar(value=0)
    get_over_summary_check_button = tk.Checkbutton(window, variable=get_over_summary_check)
    get_over_summary_check_button.place(x=130, y=90)

    get_over_summary_button = tk.Button(window, text="取得關帳彙總",
                                        command=lambda: today_file.get_file(choice.get(), "over_summary"))
    get_over_summary_button.place(x=160, y=90)


    window.mainloop()

