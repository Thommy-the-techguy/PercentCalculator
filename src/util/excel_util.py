import pandas as ps
import calendar

def get_last_sheet_name(file_name: str) -> str:
    excel_file = ps.ExcelFile(file_name)

    return excel_file.sheet_names[-1]

def read_excel_data(file_name: str) -> dict:
    """
    Reads last added excel sheet
    """

    return dict(ps.read_excel(file_name, sheet_name=get_last_sheet_name(file_name)))

def form_new_data_frame(data: dict, indicies: (list | None), days_value: (int | None)=None) -> ps.DataFrame:
    new_data: dict = {}
    
    if indicies == None:
        indicies = [i for i in range(0, len(data[list(data.keys())[0]]))]

    for key in data.keys():
        new_data[key] = []

    for key in data.keys():
        new_index: int = 0
        for index in indicies:
            match key:
                case "№ п/п":
                    new_data[key].append(new_index + 1)
                case "Кол-во дней":
                    if days_value != None:
                        new_data[key].append(days_value)
                    else:
                        new_data[key].append(data[key][index])
                case "Остаток неоплаченной суммы":
                    sum_ttn: float = float(data["Сумма ТТН"][index])
                    paid_sum: float = float(data["Оплаченная сумма"][index])

                    new_data[key].append(sum_ttn - paid_sum)    
                case "Сумма коммерческого займа":
                    sum_left: float = float(new_data["Остаток неоплаченной суммы"][new_index])
                    percents: float = float(data["%-я ставка"][index]) / 100
                    days_amount: int = int(new_data["Кол-во дней"][new_index])
                        
                    new_data[key].append(round(sum_left * percents * days_amount, 2))
                case _:
                    new_data[key].append(data[key][index])

            new_index += 1
        
    return ps.DataFrame(new_data)

def write_data_to_excel(data_frame: ps.DataFrame, file_path: str, sheet_name: str):
    with ps.ExcelWriter(file_path, date_format="%d.%m.%Y", mode="a") as writer:
        workbook = writer.book
        try:
            workbook.remove(workbook[sheet_name])
        except:
            print("Worksheet doesn't exist!")
        finally:
            data_frame.to_excel(writer, sheet_name=sheet_name, index=False)

def merge_frames(data: dict, additional_data: dict, position: (int | None)=None) -> ps.DataFrame:
        new_data: dict = create_dict_with_frame_keys(data)

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

def create_dict_with_frame_keys(template: dict) -> dict:
        output: dict = {}

        for key in template.keys():
            output[key] = []

        return output