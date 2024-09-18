import json
from typing import final

@final
class JsonUtil():
    def __init__(self) -> None:
        pass

    @staticmethod
    def serialize_current_date(date_str: str, file_name: str):
        with open(file_name, "w") as file:
            json.dump(date_str, file)

    @staticmethod
    def deserialize_current_date(file_name: str) -> object:
        result: object = None

        with open(file_name, "r") as file:
            result = json.load(file)

        return result  
