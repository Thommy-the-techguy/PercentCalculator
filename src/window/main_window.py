from tkinter import *
from tkcalendar import Calendar
from typing import final
from util.json_util import JsonUtil
import pandas as ps
import calendar
from window.payment_window import PaymentWindow
from util.excel_util import *

@final
class MainWindow(Tk): #TODO: remove unnecessary code, since inheritance can be used instead
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
            PaymentWindow(self.__main_window)

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
    
    # def __form_new_data_frame(self, data: dict, indicies: list) -> ps.DataFrame:
    #     new_data: dict = {}

    #     for key in data.keys():
    #         new_data[key] = []

    #     for key in data.keys():
    #         for index in indicies:
    #             match key:
    #                 case "Кол-во дней":
    #                     selected_year: int = int(self.get_date_picker().selection_get().year)
    #                     selected_month: int = int(self.get_date_picker().selection_get().month)

    #                     new_data[key].append(calendar.monthrange(selected_year, selected_month)[1])
    #                 case "Остаток неоплаченной суммы":
    #                     sum_ttn: float = float(new_data["Сумма ТТН"][index])
    #                     paid_sum: float = float(new_data["Оплаченная сума"][index])

    #                     new_data[key].append(sum_ttn - paid_sum)    
    #                 case "Сумма коммерческого займа":
    #                     sum_left: float = float(new_data["Остаток неоплаченной суммы"][index])
    #                     percents: float = float(new_data["%-я ставка"][index]) / 100
    #                     days_amount: int = int(new_data["Кол-во дней"][index])
                        
    #                     new_data[key].append(round(sum_left * percents * days_amount, 2))
    #                 case _:
    #                     new_data[key].append(data[key][index])
        
    #     return ps.DataFrame(new_data)
    def __days_func(self) -> int:
        year: int = self.get_date_picker().selection_get().year
        month: int = self.get_date_picker().selection_get().month

        return calendar.monthrange(year, month)[1]
    
    # def __write_data_to_excel(self, data_frame: ps.DataFrame, file_path: str):
    #     month_chosen = int(self.get_date_picker().get_displayed_month()[0]) 

    #     with ps.ExcelWriter(file_path, date_format="%d.%m.%Y", mode="a") as writer:
    #         data_frame.to_excel(writer, sheet_name=f"{month_chosen}", index=False)
