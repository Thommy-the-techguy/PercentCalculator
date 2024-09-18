import datetime
from util.json_util import JsonUtil
from window.main_window import MainWindow
from tkinter import *
from tkcalendar import Calendar
import os

def check_if_new_month():
    pass

if __name__ == '__main__':
    if not os.path.exists("ser.json"):
        JsonUtil.serialize_current_date(str(datetime.date.today().month), "ser.json")
    else:
        print("File ser.json exists!")

    main_window: MainWindow = MainWindow()
    main_window.get_main_window().mainloop()
    date_picker: Calendar = main_window.get_date_picker()
