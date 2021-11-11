import line
import json
from datetime import datetime

with open("setup.json", encoding="utf-8") as f:
    setup = json.load(f)

path = setup["path"]
fmt = setup["fmt"]
start = datetime.strptime(setup["start"], fmt)
end = datetime.strptime(setup["end"], fmt)
name = setup["name"]
text = setup["text"]

chat = line.Chat.read(path).select(start, end, name, text)
chat.save("result_message.html", reset_index=True)