import re
from datetime import datetime
import pandas as pd
import csv, json
from copy import deepcopy, copy


class Text:
    def __init__(self, *args: str):
        '''args: string'''
        L = list()
        for i in args:
            L.extend(i.removesuffix("\n").split(sep="\n"))
        self.list: list = L

    def __iter__(self):
        for i in self.list:
            yield i

    def __repr__(self):
        return f"line.Text({', '.join(self.list)})"

    def __str__(self):
        return Text.str(self, "\n")

    def __iadd__(self, text):
        if type(text) is str:
            text = Text(text)
        if type(text) is Text:
            self.list.extend(text.list)
        return self

    def str(self, delimiter: str = "\n") -> str:
        return delimiter.join(self.list)


class Message:
    def __init__(self, time=None, name=None, text=None, fmt=None, **kwargs):
        '''\
            time: str or pd.Timestamp or datetime\n
            name: str\n
            text: str or list or Text\n
            fmt: string format for time\n
            kwargs: any additional attributes'''
        if type(time) is str:
            self.time: datetime = datetime.strptime(time, fmt)
        elif type(time) is pd.Timestamp:
            self.time = time.to_pydatetime()
        elif type(time) is datetime:
            self.time = time
        if type(name) is str:
            self.name: str = name
        if type(text) is str:
            self.text: Text = Text(text)
        elif type(text) is list:
            self.text = Text(*text)
        elif type(text) is Text:
            self.text = text
        for i in kwargs:
            setattr(self, i, kwargs[i])

    def __iter__(self):
        for i in vars(self):
            yield i

    def __repr__(self):
        return f"line.Message({str(self.__dict__)})"

    def to_dict(self) -> dict:
        return vars(self)

    def from_Series(series: pd.Series):
        '''series: pd.Series'''
        return Message(**(series.to_dict()))

    def to_Series(self) -> pd.Series:
        '''return pd.Series'''
        return pd.Series(self.to_dict())


class Chat:
    def __init__(self, message, **kwargs):
        '''\
            message: list of Message or list of DataFrame or DataFrame\n
            kwargs: any additional attributes'''
        if type(message) is list:
            if type(message[0]) is Message:
                df = pd.DataFrame(map(lambda x: x.to_dict(), message))
            elif type(message[0]) is pd.DataFrame:
                df = pd.concat(message)
            if "time" in df:
                df = df.sort_values("time", ignore_index=True)
        elif type(message) is pd.DataFrame:
            df = message
        self.message = df
        for i in kwargs:
            setattr(self, i, kwargs[i])

    def __repr__(self):
        return f"<line.Chat object>"

    def append(self, chat, reset_index: False):
        '''\
            combine Chat object with another Chat object\n
            chat: Chat\n
            reset_index'''
        df = pd.concat([self.message, chat.message])
        df = df.drop_duplicates()
        if reset_index:
            df = df.reset_index(drop=True)
        return Chat(df)

    def select(self, time_start: datetime = None, time_end: datetime = None, name: list = None, text: list[str] = None, fmt=None, reset_index=False): # yapf: disable
        '''\
            return Chat object\n
            time_start, time_end: datetime or str\n
            name:list of str\n
            text:list of str with regular expression\n
            fmt: string format for time\n
            reset_index'''
        c = copy(self)
        df = c.message
        if time_start:
            if type(time_start) is str:
                time_start = datetime.strptime(time_start, fmt)
            if type(time_start) is datetime:
                df = df[df["time"] >= time_start]
        if time_end:
            if type(time_end) is str:
                time_end = datetime.strptime(time_end, fmt)
            if type(time_start) is datetime:
                df = df[df["time"] < time_end]
        if name:
            df = df[df["name"].isin(name)]
        if text:
            df = df[df["text"].astype(str).str.contains('|'.join(text))]
        if reset_index:
            df = df.reset_index(drop=True)
        c.message = df
        return c

    def read(path: str):
        '''\
            Supported file type: txt, csv, json'''
        ext = path.rsplit(".", maxsplit=1)[-1]
        if ext == "txt":
            L = list()
            with open(path, encoding="utf-8-sig") as f:
                chatname = f.readline().removeprefix("[LINE] ").removesuffix("的聊天記錄\n")
                save_time = datetime.strptime(f.readline().removeprefix("儲存日期：").removesuffix("\n"), "%Y/%m/%d %H:%M")
                for ln in f:
                    if re.match(r"^\d{2}:\d{2}\t.+\t.+$", ln):
                        i = ln.split("\t")
                        L.append(Message(time=f"{date} {i[0]}", name=i[1], text=i[2], fmt="%Y/%m/%d %H:%M"))
                    elif re.match(r"^\d{2}:\d{2}\t.+$", ln):
                        i = ln.split("\t")
                        L.append(Message(time=f"{date} {i[0]}", name="LINE", text=i[1], fmt="%Y/%m/%d %H:%M"))
                    elif re.match(r"^\d{4}/\d{1,2}/\d{1,2}", ln):
                        date = re.match(r"^\d{4}/\d{1,2}/\d{1,2}", ln).group(0)
                    elif ln == "\n":
                        continue
                    else:
                        L[-1].text += ln
            return Chat(L, chat=chatname, save_time=save_time)
        elif ext == "csv":
            df = pd.read_csv(path, index_col=0)
            df["text"] = df["text"].apply(lambda x: Text(x))
            return Chat(df)
        elif ext == "json":
            df = pd.read_json(path, orient="index")
            df["text"] = df["text"].apply(lambda x: Text(*x))
            return Chat(df)
        else:
            raise ValueError("Unsupported file extension")

    def save(self, path: str, reset_index=False):
        '''\
            Supported file type: xlsx,csv, json, html\n
            reset_index
            '''
        ext = path.rsplit(".", maxsplit=1)[-1]
        df = copy(self.message)
        if reset_index:
            df = df.reset_index(drop=True)
        if ext == "xlsx":
            with pd.ExcelWriter(path, engine="xlsxwriter", engine_kwargs={'options': {'strings_to_formulas': False }}) as writer:# yapf: disable
                df.to_excel(writer)
            return None
        elif ext == "csv":
            df["text"] = df['text'].astype(str)
            df.replace(r'\n', r'\\n', regex=True).to_csv(path)
            return None
        elif ext == "json":
            df["text"] = df["text"].apply(lambda x: x.list)
            df.to_json(path, orient="index", date_format="iso", force_ascii=False, indent=4)
            return None
        elif ext == "html":
            df["text"] = df['text'].astype(str)
            html_table = df.replace(r"\n", "<br/>", regex=True).to_html(escape=False)
            with open(path, mode='w', encoding="utf-8") as f:
                f.write(f'<html>\n<head><meta charset="utf-8"></head>\n<body>\n{html_table}\n</body>\n</html>')
            return None
        else:
            raise ValueError("Unsupported file extension")
