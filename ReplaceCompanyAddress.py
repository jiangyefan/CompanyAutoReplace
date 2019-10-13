#!/usr/bin/env python3

from openpyxl import load_workbook
import cpca


class ReplaceCompanyAddress:
    def __init__(self,filepath):
        self.filepath = filepath

    def read(self):
        # 打开excel表格
        wb = load_workbook(self.filepath)

        # 获取sheet：
        table = wb.get_sheet_names()  # 通过表名获取

        return 0



testaddress="C:\\Users\\jiangjiang\\Desktop\\work\\部门最新.xlsx"
test=ReplaceCompanyAddress(testaddress)
test.read()


address=["浙江省杭州市学院路77号黄龙国际中心商场 4层431号 海马体照相馆".replace(" ","")]
seg_list = cpca.transform(address, cut=False)
print(seg_list)