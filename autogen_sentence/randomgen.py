import sys
import random
import argparse
import pandas as pd
import openpyxl
import os

from openpyxl.workbook import Workbook
def randomgenarticles(path):
    #wb=openpyxl.load_workbook("C:\\Users\\Administrator\\Desktop\\本地大模型\\爆款文案批量生成器\\爆款文案_2024.xlsx")
    #path = "爆款文案_2024_新文案_句级_1.xlsx"
    wb=openpyxl.load_workbook(path)
    sheet=wb.active
    # wb_new = Workbook()
    # sheet_new = wb_new.active
    #print(sheet.max_column)
    rows = [[""] for i in range(sheet.max_column+1)]
    #print(len(rows))
    offset = 8
    for col in range(1, sheet.max_column + 1):
        for row in range(offset, sheet.max_row + 1):
            v = sheet.cell(row=row, column=col).value
            rows[col].append(v)
            #print(v)
        #print("next col")
    dest = 10000
    articles = []
    mx_c = sheet.max_column
    mx_r = sheet.max_row-offset+1
    for i in range(dest):
        article = ""
        for j in range(1, mx_c+1):
            idx = random.randint(1, mx_r)
            article = article+rows[j][idx]
        print(str(i)+":"+article)
        articles.append(article)
    #print(articles)
    wb_new = Workbook()
    sheet_new = wb_new.active
    for i in range(len(articles)):
        cell = sheet_new.cell(row=i + 1, column=1)
        cell.value = articles[i]
    wb_new.save(path[:-5] +'_random.xlsx')


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--path", type=str, default="/home/liuge/Desktop/本地大模型/爆款文案_2024_新文案_句级_1.xlsx")
    args = parser.parse_args()
    randomgenarticles(args.path)
