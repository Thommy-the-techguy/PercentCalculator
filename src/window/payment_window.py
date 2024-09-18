from typing import final
from tkinter import *
from util.excel_util import *
import pandas as ps
import datetime

@final
class PaymentWindow():
    __window: Toplevel = None
    __ttn_label: Label = None
    __payment_sum_label: Label = None
    __ttn_text_box: Text = None
    __payment_sum_text_box: Text = None
    __payment_button: Button = None

    def __init__(self, master: Tk) -> None:
        self.__window = Toplevel(master)
        self.__setup_ttn_label()
        self.__setup_ttn_text_box()
        self.__setup_payment_sum_label()
        self.__setup_payment_sum_text_box()
        self.__setup_payment_button()

    def __setup_ttn_label(self):
        self.__ttn_label = Label(self.__window, text="ТТН:")
        self.__ttn_label.pack(anchor="nw")

    def __setup_ttn_text_box(self):
        self.__ttn_text_box: Text = Text(self.__window, height=1, width=20)
        self.__ttn_text_box.pack(anchor="w")

    def __setup_payment_sum_label(self):
        self.__payment_sum_label: Label = Label(self.__window, text="Сумма оплаты:")
        self.__payment_sum_label.pack(anchor="w")

    def __setup_payment_sum_text_box(self):
        self.__payment_sum_text_box: Text = Text(self.__window, height=1, width=20)
        self.__payment_sum_text_box.pack(anchor="w")

    def __setup_payment_button(self):
        self.__payment_button: Button = Button(self.__window, text="Внести оплату", command=lambda: self.__payment_button_action())
        self.__payment_button.pack(anchor="s")

    def __payment_button_action(self):
        data: dict = self.__add_payment()
        data_frame: ps.DataFrame = form_new_data_frame(data, None, None)
        file_name: str = "data.xlsx"

        write_data_to_excel(data_frame, file_name, get_last_sheet_name(file_name))

    def __add_payment(self) -> dict:
        ttn: int = int(self.__ttn_text_box.get("1.0", END).strip())
        payment_sum: float = float(self.__payment_sum_text_box.get("1.0", END).strip())
        excel_data: dict = read_excel_data("data.xlsx")

        item_index: int = 0
        for value in excel_data["№ ТТН"]:
            if value == ttn:
                if excel_data["Оплаченная сумма"][item_index] == 0:
                    new_days_amount: int = datetime.date.today().day
                    excel_data["Кол-во дней"][item_index] = new_days_amount
                else:
                    prev_payment_day: int = datetime.datetime.strptime(str(excel_data["Дата оплаты"][item_index]), "%Y-%m-%d %H:%M:%S").date().day
                    new_days_amount: int = datetime.date.today().day - prev_payment_day

                    excel_data["Кол-во дней"][item_index] = new_days_amount
                
                excel_data["Оплаченная сумма"][item_index] += payment_sum

                #TODO: remove datetime later and replace with date from datepicker, if no more ideas will come
                excel_data["Дата оплаты"][item_index] = datetime.date.today().strftime("%d.%m.%Y")

            item_index += 1  

        return excel_data

    def get_window(self) -> Toplevel:
        return self.__window
