import yaml
import os

class Setting:
    def __init__(self, path: str):
        self.checkpath(path)
        self.path = path
        self._settings = {}
        self.load()
    
    def checkpath(self, path: str):
        if not os.path.isfile(path):
            data = {
                "listProduct": [
                    {
                    "name": "Cal Stearate",
                    "key": [
                        "calcium"
                    ]
                    },
                    {
                    "name": "Zinc Stearate",
                    "key": [
                        "zinc"
                    ]
                    }
                ],
                "listExcludeName": [
                    "company",
                    "limited",
                    "ltd",
                    "tradding",
                    "jsc",
                    "international",
                    "corp",
                    "corporation",
                    "joint stock company",
                    "pte",
                    "ooo"
                ],
                "weightUnit": [
                    {
                    "name": "Kilogram",
                    "exchange": 1,
                    "key": [
                        "kg",
                        "kilogram",
                        "kilograms"
                    ]
                    },
                    {
                    "name": "Ton",
                    "exchange": 1000,
                    "key": [
                        "ton",
                        "tonne",
                        "tons",
                        "tonnes"
                    ]
                    },
                    {
                    "name": "Gram",
                    "exchange": 0.001,
                    "key": [
                        "g",
                        "gram",
                        "grams"
                    ]
                    }
                ]
            }
            with open(path, 'w', encoding='utf-8') as file:
                yaml.safe_dump(data, file)

    def load(self):
        with open(self.path, 'r', encoding='utf-8') as file:
            self._settings = yaml.safe_load(file)

    def get(self, key: str, default=None):
        return self._settings.get(key, default)

    def set(self, key: str, value):
        self._settings[key] = value
        self.save()

    def save(self):
        with open(self.path, 'w', encoding='utf-8') as file:
            yaml.safe_dump(self._settings, file)