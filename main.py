import pandas as pd
from docxtpl import DocxTemplate
import jinja2
import json
from random import randrange
import os
import sys

cwd = os.getcwd()
how_to_use_command = """$ python main.py <class_name>
- class_name = [B19, B29, 125, 234, 302]"""

if __name__ == "__main__":
    class_room_list = ["B19", "B29", "125", "234", "302"]
    class_name = {
        "B19": "資電B19",
        "B29": "資電B29",
        "125": "資電125",
        "234": "資電234",
        "302": "土302"
    }
    if len(sys.argv) < 2:
        print("請輸入教室名稱")
        print(how_to_use_command)
        sys.exit()
    class_room_name = sys.argv[1]
    if class_room_name not in class_room_list:
        print("教室名稱輸入錯誤")
        print(how_to_use_command)
        sys.exit()
    xlsx_name = ""
    print(cwd)
    print("read xlsx")
    for xlsx_file in os.listdir():
        if xlsx_file.endswith(".xlsx"):
            xlsx_name = xlsx_file
            print(xlsx_file)
            break
    df = pd.read_excel(xlsx_name, sheet_name="成績")
    name_list = df["姓名"].tolist()
    name_len = len(name_list)
    doc = DocxTemplate(os.path.join(cwd, "template_{}.docx".format(class_room_name)))
    json_var = json.loads(open("var.json", "r", encoding="UTF-8").read())
    for i in range(1, name_len+1):
        list_len = len(name_list)
        index = "var%d"%i
        if list_len == 1:
            json_var[index] = name_list[0]
        else:
            rand_index = randrange(1, len(name_list), 1)
            json_var[index] = name_list.pop(rand_index)

    context = json_var
    jinja_env = jinja2.Environment()
    doc.render(context, jinja_env)
    print("save")
    doc.save(os.path.join(cwd, "{}座位表.docx".format(class_name[class_room_name])))
