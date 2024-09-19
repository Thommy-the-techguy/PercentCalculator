from typing import final
from tkinter import *
from util.excel_util import *
import pandas as ps
import datetime
from tkcalendar import Calendar

@final
class PaymentWindow():
    __window: Toplevel = None
    __ttn_label: Label = None
    __payment_sum_label: Label = None
    __ttn_text_box: Text = None
    __payment_sum_text_box: Text = None
    __payment_button: Button = None
    __master_window = None

    def __init__(self, master) -> None:
        self.__window = Toplevel(master.get_main_window())
        self.__master_window = master
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
        excel_data: dict = read_excel_data("data.xlsx")
        additional_data, item_index = self.__add_payment()
        data_frame: ps.DataFrame = self.merge_frames(excel_data, additional_data, item_index)
        file_name: str = "data.xlsx"

        write_data_to_excel(data_frame, file_name, get_last_sheet_name(file_name))

    def __add_payment(self) -> tuple:
        ttn: int = int(self.__ttn_text_box.get("1.0", END).strip())
        payment_sum: float = float(self.__payment_sum_text_box.get("1.0", END).strip())
        excel_data: dict = read_excel_data("data.xlsx")
        additional: dict = self.create_dict_with_frame_keys(excel_data)

        item_index: int = 0
        for index in range(0, len(excel_data["№ ТТН"])):
            if excel_data["№ ТТН"][index] == ttn and ((index + 1) >= len(excel_data["№ ТТН"]) or (excel_data["№ ТТН"][index + 1] != ttn)):
                master_window_calendar: Calendar = self.__master_window.get_date_picker()
                calendar_date: datetime.date = datetime.datetime.strptime(master_window_calendar.get_date(), "%d.%m.%Y").date()
                calendar_day: int = calendar_date.day

                for key in additional.keys():
                    match key:
                        case "№ п/п":
                            additional["№ п/п"].append(excel_data["№ п/п"][item_index] + 1)
                        case "Оплаченная сумма":
                            if excel_data["Оплаченная сумма"][item_index] == 0:
                                new_days_amount: int = calendar_day

                                additional["Кол-во дней"].append(new_days_amount)
                            else:
                                prev_payment_day: int = datetime.datetime.strptime(str(excel_data["Дата оплаты"][item_index]), "%d.%m.%Y").date().day
                                new_days_amount: int = calendar_day - prev_payment_day

                                additional["Кол-во дней"].append(new_days_amount)
                            
                            additional["Оплаченная сумма"].append(payment_sum + excel_data["Оплаченная сумма"][item_index])
                        case "Дата оплаты":
                            additional["Дата оплаты"].append(calendar_date.strftime("%d.%m.%Y")) 
                        case "Остаток неоплаченной суммы":
                            sum_ttn: float = additional["Сумма ТТН"][0]
                            paid_sum: float = additional["Оплаченная сумма"][0]

                            additional["Остаток неоплаченной суммы"].append(sum_ttn - paid_sum)
                        case "Сумма коммерческого займа":
                            left_to_pay: float = additional["Остаток неоплаченной суммы"][0]
                            percentage: float = additional["%-я ставка"][0] / 100.0
                            days: int = additional["Кол-во дней"][0]

                            additional["Сумма коммерческого займа"].append(round(left_to_pay * percentage * days, 2))
                        case "Кол-во дней":
                            pass
                        case _:
                            additional[key].append(excel_data[key][item_index])

                break

            item_index += 1  

        return (additional, item_index + 1)
    
    def merge_frames(self, data: dict, additional_data: dict, position: (int | None)=None) -> ps.DataFrame:
        new_data: dict = self.create_dict_with_frame_keys(data)

        for key in data.keys():
            for value in data[key]:
                new_data[key].append(value)

        for key in new_data.keys():
            for value in additional_data[key]:
                if position == None:
                    new_data[key].append(value)
                else:
                    new_data[key].insert(position, value)
                    if key == "№ п/п":
                        for i in range(position + 1, len(new_data[key])):
                            new_data[key][i] += 1

        return ps.DataFrame(new_data)

    def create_dict_with_frame_keys(self, template: dict) -> dict:
        output: dict = {}

        for key in template.keys():
            output[key] = []

        return output

    def get_window(self) -> Toplevel:
        return self.__window
