import json


class Assistant:
    def __init__(self, flName='Design_info.json'):
        self.fileName = flName

    def get_variable_value(self, variable):
        with open(self.fileName) as jsonFile:
            dictionary = json.load(jsonFile)
            value = dictionary[variable]
        return value

    def set_variable_value(self, variable, value):
        with open(self.fileName) as jsonFile:
            dictionary = json.load(jsonFile)
        with open(self.fileName, 'w') as jsonFile:
            dictionary[variable] = value
            json.dump(dictionary, jsonFile)

    def set_default_variable(self, key, value):
        with open(self.fileName) as jsonFile:
            dictionary = json.load(jsonFile)
        dictionary.setdefault(str(key), value)
        with open(self.fileName, 'w') as jsonFile:
            json.dump(dictionary, jsonFile)

    def reset_values(self):
        default_values = {
            "Design_info":
                {
                    "": ""
                }
        }
        dictionary = default_values.get(self.fileName)
        with open(self.fileName, 'w') as jsonFile:
            json.dump(dictionary, jsonFile)


if __name__ == '__main__':
    ast = Assistant()
