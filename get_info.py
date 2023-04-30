import requests


class GetInfo:
    def __init__(self):
        session = requests.Session()
        response = session.post("https://miet.ru/schedule/groups")
        index = response.text.find("document.cookie=\"")
        index += len("document.cookie=\"")
        cookie = ""
        while response.text[index] != "\"":
            cookie += response.text[index]
            index += 1
        session.headers.update({"cookie": cookie})
        self.session = session

    def get_groups(self):
        response = self.session.post("https://miet.ru/schedule/groups")
        return response.json()

    def get_schedule(self, group_name: str):
        response = self.session.post("https://miet.ru/schedule/data", {"group": group_name})
        return response.json()
