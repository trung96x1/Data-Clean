import yaml

class Setting:
    def __init__(self, path: str):
        self.path = path
        self._settings = {}
        self.load()

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