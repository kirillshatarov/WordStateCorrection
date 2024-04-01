import asyncio
import os
import json


class FileReader:
    def __init__(self, name: str, ):
        self.name = name

    async def read_file(self):
        with open(f"../files/gost/{self.name}",encoding='UTF-8') as f:
            return json.load(f)

    @classmethod
    def get_files(cls):
        return {gost.replace('.json', ''): gost for gost in os.listdir('../files/gost/')}

if __name__ == '__main__':
    js_dict =  FileReader('research.json')
    file_cont = asyncio.run(js_dict.read_file())
    print(file_cont)
    print(js_dict.get_files())

