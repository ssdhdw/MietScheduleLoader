import requests
import asyncio
import aiohttp
import sqlite3
from tqdm import tqdm

response = requests.get("https://miet.ru/")
headers = { "Cookie" : response.text.split('"')[1].split(";")[0]}

async def get_group(session, group, pbar):
    async with session.post("https://miet.ru/schedule/data", headers = headers, data = {"group": group}) as r:
        pbar.update(1)
        return await r.json()

async def get_all_groups(session):
    groups = requests.post("https://miet.ru/schedule/groups", headers=headers).json()
    tasks = []
    with tqdm(total = len(groups), unit=" group") as pbar:
        for group in groups:
            task = asyncio.create_task(get_group(session, group, pbar))
            tasks.append(task)
        result = await asyncio.gather(*tasks)
    return result
    

async def main():
    async with aiohttp.ClientSession() as session:
        groups = await get_all_groups(session)
    con = sqlite3.connect("data.db")
    cur = con.cursor()
    cur.execute(
        """CREATE TABLE groups (
        name TEXT,
        day INTEGER,
        week INTEGER,
        time INTEGER,
        room TEXT,
        subject TEXT,
        teacher TEXT
        );""")
    cur.execute(
        """CREATE TABLE rooms (
        name TEXT,
        code INTEGER,
        day INTEGER,
        week INTEGER,
        time INTEGER,
        group_name TEXT
        );""")
    print("Data downloaded")
    for group in tqdm(groups):
        group_values = []
        rooms_value = []
        for i in group["Data"]:
            group_values.append((
                i["Group"]["Name"],
                i["Day"],
                i["DayNumber"],
                i["Time"]["Code"],
                i["Room"]["Name"],
                i["Class"]["Name"],
                i["Class"]["TeacherFull"]
            ))
            rooms_value.append((
                i["Room"]["Name"],
                i["Room"]["Code"],
                i["Day"],
                i["DayNumber"],
                i["Time"]["Code"],
                i["Group"]["Name"]
            ))
        cur.executemany("INSERT INTO groups VALUES(?,?,?,?,?,?,?);", group_values)
        cur.executemany("INSERT INTO rooms VALUES(?,?,?,?,?,?);", rooms_value)
    cur.close()
    con.commit()
    con.close()

if __name__ == "__main__":
    asyncio.run(main())