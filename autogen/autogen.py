import pandas as pd
import openpyxl
import os

from openpyxl.workbook import Workbook
from transformers import AutoTokenizer, AutoModel

os.environ['CUDA_VISIBLE_DEVICES'] = '1'
MODEL_PATH = os.environ.get('MODEL_PATH', 'THUDM/chatglm3-6b-128k')
TOKENIZER_PATH = os.environ.get("TOKENIZER_PATH", MODEL_PATH)

tokenizer = AutoTokenizer.from_pretrained(TOKENIZER_PATH, trust_remote_code=True)
model = AutoModel.from_pretrained(MODEL_PATH, trust_remote_code=True, device_map="auto").eval()

wb=openpyxl.load_workbook("C:\\Users\\Administrator\\Desktop\\本地大模型\\爆款文案批量生成器\\爆款文案_2024.xlsx")
sheet=wb["爆款文案_2024"]
cnt = sheet['B2'].value
if not cnt:
    cnt = int(cnt)
print("将每条自动生成"+str(cnt)+"条")
for row in range(10,sheet.max_row + 1):
    cellS=sheet.cell(row=row,column=2)
    if cellS.value:
        print(cellS.value)
    query = "我给你发一段话，你保持原来的意思不变，帮我去润色一下啊，润色的长度和原来的差不多。 下面是你要润色的文本："+cellS.value
    wb_new = Workbook()
    sheet_new = wb_new.active
    sheet_new.title = str(row)
    past_key_values, history = None, []
    for i in range(cnt):
        current_length = 0
        article = ""
        for response, history, past_key_values in model.stream_chat(tokenizer, query, history=history, top_p=1,
                                                                    temperature=0.8,
                                                                    past_key_values=past_key_values,
                                                                    return_past_key_values=True):
            print(response[current_length:], end="", flush=True)
            current_length = len(response)
            article = response
        sheet_new['A'+str(i+1)] = article
        print()
        wb_new.save('C:\\Users\\Administrator\\Desktop\\本地大模型\\爆款文案批量生成器\\爆款文案_2024_新文案_'+str(row-9)+'.xlsx')