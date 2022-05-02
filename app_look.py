import numpy as np
import xlwings as xw


def app():
    wb = xw.books.active
    sheet = wb.sheets.active

    shps = sheet.shapes

    for s in sheet.shapes:
        print(f"-- {s.name} --")
        print(f"type: {s.type}")
        print(f"top: {s.top}")      # 垂直位置
        print(f"left: {s.left}")    # 水平位置
        print(f"width: {s.width}")  # 幅
        print(f"height: {s.height}")# 高さ
        print(f"parent: {s.parent}")# 親オブジェクト

if __name__=="__main__":
    app()