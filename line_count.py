import line
import json
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta

with open("setup.json", encoding="utf-8") as f:
    setup = json.load(f)

path = setup["path"]
fmt = setup["fmt"]
start = datetime.strptime(setup["start"], fmt)
end = datetime.strptime(setup["end"], fmt)
delta = relativedelta(**setup["delta"])
name = setup["name"]
text = setup["text"]

chat = line.Chat.read(path).select(start, end, name, text)
time = start
L = list()
while start <= time < end:
    count = chat.select(time, time + delta).message["name"].value_counts()
    count.name = time
    L.append(count)
    time += delta
df = pd.concat(L, axis=1)
df = df.fillna(0)
df = df.astype(int)
df = df.rename_axis("Name")
df.insert(0, "Total", df.sum(axis=1))
df = df.append(df.sum(axis=0).rename("Total"))
df = df.sort_values("Total", ascending=False)
df.insert(0, "Rank", range(len(df)))

with pd.ExcelWriter("result_count.xlsx", engine="xlsxwriter", datetime_format="m/d") as writer:
    df.to_excel(writer)
    wb = writer.book
    ws = wb.worksheets()[0]
    (row, col) = df.shape
    ws.conditional_format(2, 3, row, col, {"type": "2_color_scale", "min_color": "#ffffff", "max_color": "#ff0000"})
    ws.conditional_format(1, 3, 1, col, {"type": "2_color_scale", "min_color": "#ffffff", "max_color": "#0070c0"})
    ws.conditional_format(2, 2, row, 2, {"type": "data_bar", "bar_color": "#0070c0", "bar_solid": True})
    ws.autofilter(0, 0, row, col)