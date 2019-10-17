#!/usr/bin/env python3

from openpyxl import load_workbook
import cpca
import re
from openpyxl.styles import Font, colors, Alignment
import json



class ReplaceCompanyAddress:
    def __init__(self,filepath,parameter,sendAddress,recAddress):
        self.filepath = filepath
        self.parameter = parameter
        self.sendAddress = sendAddress
        self.recAddress = recAddress

    # 先判断有没有门店一列
    def setUp(self):
        wb = load_workbook(self.filepath)
        ws = wb.get_sheet_by_name(self.parameter)
        addressInfo=ws.cell(row=1,column=ws.max_column).value
        try:
            if re.search("门店",addressInfo).group():
                return 1
        except:
            return 0



    # 处理门店数据
    def addressInfoProcess(self):
        wb = load_workbook(self.filepath)
        ws = wb.get_sheet_by_name(self.parameter)
        #添加门店title
        ws[ws.cell(row=1,column=ws.max_column+1).coordinate]="门店"
        # ws[ws.cell(row=1, column=ws.max_column).coordinate].alignment = Alignment(horizontal='center', vertical='center')

        #提取所有行
        rows=[]
        for row in ws.iter_rows():
            rows.append(row)

        #判断寄件地址和收件地址的列
        for addressCoordinate in range(len(rows[0])):
            if self.sendAddress == rows[0][addressCoordinate].value:
                sendAddressCol = rows[0][addressCoordinate].column
            elif self.recAddress == rows[0][addressCoordinate].value:
                recAddressCol = rows[0][addressCoordinate].column

        # 地址转换门店
        # 先排除本公司的地址
        ruleAddressAll = ["缦图", "外包", "和达高科创新服务中心", "科技园路65号", "科技园65号"]
        with open("./Administrative-divisions-of-China", "r") as f:
            load_address=json.load(f)

        for addressInfo in range(1,len(rows)):
            for ruleAddress in ruleAddressAll:
                if ruleAddress in rows[addressInfo][sendAddressCol-1].value :
                    rows[addressInfo][ws.max_column - 1].value="行政部"
                    break
                else:
                    addressWord=rows[addressInfo][sendAddressCol - 1].value.split()
                    province=cpca.transform(addressWord, cut=False).省.values[0]
                    city=cpca.transform(addressWord, cut=False).市.values[0]
                    area=cpca.transform(addressWord, cut=False).区.values[0]
                    if province==city and area !="":
                        for realaddress in load_address[city][area]:
                            if realaddress["address"] in rows[addressInfo][sendAddressCol - 1].value:
                                rows[addressInfo][ws.max_column - 1].value = realaddress["name"]
                                break
                        break
                    elif province!=city and area !="":
                        for realaddress in load_address[province][city][area]:
                            if realaddress["address"] in rows[addressInfo][sendAddressCol - 1].value:
                                rows[addressInfo][ws.max_column - 1].value = realaddress["name"]
                                break
                        break
                    else:
                        print(rows[addressInfo][sendAddressCol - 1].coordinate+"坐标中的地址异常,地址内容为"+rows[addressInfo][sendAddressCol - 1].value)
                        break




        wb.save(testaddress)








if __name__ == '__main__':
    testaddress="/Users/jiangjiang/Desktop/顺丰.xlsx"
    test=ReplaceCompanyAddress(testaddress,"Sheet0","寄件公司地址","收件地址")
    if test.setUp()==1:
        print("已有门店信息,请检查！")
    else:
        test.addressInfoProcess()



    # address=["浙江省杭州市学院路77号黄龙国际中心商场 4层431号 海马体照相馆".replace(" ","")]
# seg_list = cpca.transform(address, cut=False)
# print(seg_list)