from tkinter import *
from tkcalendar import Calendar
from typing import final
from util.json_util import JsonUtil
import pandas as ps
import calendar
from window.payment_window import PaymentWindow
from util.excel_util import *

@final
class MainWindow():
    __main_window: Tk = None
    __date_picker: Calendar = None
    __confirm_button: Button = None
    
    def __init__(self) -> None:
        self.__main_window = self.__setup_main_window()
        self.__setup_datepicker_label()
        self.__setup_datepicker(self.__main_window)
        self.__setup_confirm_button()

    def get_main_window(self) -> Tk:
        return self.__main_window
    
    def get_date_picker(self) -> Calendar:
        return self.__date_picker
        
    def __setup_main_window(self) -> Tk:
        root = Tk()
        root.title("Калькулятор процентов")
        root.geometry("320x280")

        return root

    def __setup_datepicker_label(self):
        label = Label(text="Выберите текущий день:")
        label.pack(anchor="nw")

    def __setup_datepicker(self, root: Tk):
        self.__date_picker = Calendar(root, foreground="black", locale="ru_RU", font="50")
        self.__date_picker.pack(anchor="center", expand=True)

    def __setup_confirm_button(self):
        self.__confirm_button = Button(text="Выбрать", command=lambda: self.__confirm_button_action())
        self.__confirm_button.pack(anchor="s", expand=True)

    def __confirm_button_action(self):
        serialized_month: int = int(JsonUtil.deserialize_current_date("ser.json"))
        picked_month: int = self.get_date_picker().get_displayed_month()[0]

        if self.__check_if_new_month():
            print("performing record check...")

            data: dict = read_excel_data("data.xlsx")
            indicies: list = self.__check_if_not_covered(data)
            data_frame: ps.DataFrame = form_new_data_frame(data, indicies, self.__days_func())

            sheet_name: str = str(self.get_date_picker().get_displayed_month()[0])
            write_data_to_excel(data_frame, "data.xlsx", sheet_name)

            JsonUtil.serialize_current_date(f"{picked_month}", "ser.json")
        elif serialized_month == picked_month:
            print("going further!")
            PaymentWindow(self)

    def __check_if_new_month(self) -> bool:
        month_written = int(JsonUtil.deserialize_current_date("ser.json"))
        month_chosen = int(self.get_date_picker().get_displayed_month()[0])         

        return month_chosen > month_written
    
    # read excel data previous declaration
    
    def __check_if_not_covered(self, data: dict) -> list:
        list_of_indicies: list = []

        count: int = 0
        for value in data["Остаток неоплаченной суммы"]:
            if (float(value) > 0):
                list_of_indicies.append(count)

            count += 1
        
        return list_of_indicies
    
    def __days_func(self) -> int:
        year: int = self.get_date_picker().selection_get().year
        month: int = self.get_date_picker().selection_get().month

        return calendar.monthrange(year, month)[1]
