import line
import json
import pandas as pd
from datetime import datetime

with open("setup.json", encoding="utf-8") as f:
    setup = json.load(f)

path = setup["path"]
fmt = setup["fmt"]
start = datetime.strptime(setup["start"], fmt)
end = datetime.strptime(setup["end"], fmt)
delta = setup["delta"]
name = setup["name"]
ignore = setup["ignore"]
text = setup["text"]
excel_fmt = setup["excel_fmt"]

chat = line.Chat.read(path).select(start, end, name, text).message
chat["time"] = chat["time"].dt.to_period(freq=delta).dt.to_timestamp()
G = chat.groupby("time")
L = list()
for time, df in G:
    s = df["name"].value_counts()
    s.name = time
    L.append(s)
df = pd.concat(L, axis=1).fillna(0).astype(int)
df = df.rename_axis("Name")
if len(ignore):
    df = df.drop(ignore, errors='ignore')
df.insert(0, "Total", df.sum(axis=1))
df = df.append(df.sum(axis=0).rename("Total"))
df = df.sort_values("Total", ascending=False)
df.insert(0, "Rank", range(len(df)))

with pd.ExcelWriter("result_count.xlsx", engine="xlsxwriter", datetime_format=excel_fmt) as writer:
    df.to_excel(writer)
    wb = writer.book
    ws = wb.worksheets()[0]
    (row, col) = df.shape
    ws.conditional_format(2, 3, row, col, {"type": "2_color_scale", "min_color": "#ffffff", "max_color": "#ff0000"})
    ws.conditional_format(1, 3, 1, col, {"type": "2_color_scale", "min_color": "#ffffff", "max_color": "#0070c0"})
    ws.conditional_format(2, 2, row, 2, {"type": "data_bar", "bar_color": "#0070c0", "bar_solid": True})