import os
import json


class FileReader:
    def __init__(self, name: str, ):
        self.name = name

    def read_file(self):
        with open(f"../files/gost/{self.name}") as f:
            js_dict = json.load(f)
        return js_dict

    @classmethod
    def get_files(cls):
        return {gost.replace('.json', ''): gost for gost in os.listdir('../files/gost/')}

