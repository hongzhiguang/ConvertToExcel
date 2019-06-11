from util.ExcelUtil import *
from config.ProjVar import *
from data.data import *
from action.Reset import clear_cells
from openpyxl import Workbook

if not os.path.exists(excel_path):
    wb = Workbook()
    wb.save(excel_path)

print(excel_path)
excel = ParseExcel(excel_path)
clear_cells(excel)


# 遍历第一层："code":0,"msg":"Success.","data":{}
for k1,v1 in response.items():
    if  k1 == "data":
        # 遍历第二层：判断是data对应的值是否是字典型，开始遍历
        if isinstance(v1,dict):
            row = 1
            for k2,v2 in response[k1].items():
                if isinstance(v2,dict):
                    pass
                elif isinstance(v2,list):
                    for i in response[k1][k2]:
                        if isinstance(i,dict):
                            col = 1
                            for k3,v3 in i.items():
                                in_row = 0
                                if isinstance(v3,str):
                                    # pass
                                    print(row, col, v3)
                                    excel.write_cell(row, col, v3)
                                    col += 1
                                    in_row = 1
                                elif isinstance(v3,dict):
                                    pass
                                elif isinstance(v3,list):
                                    for j in v3:
                                        in_row += 1
                                        for k4,v4 in enumerate(j.values()):
                                            # pass
                                            if str(v4) == "-":
                                                v4 = "N/A"
                                            # print(row+in_row-1,k4+col,str(v4))
                                            excel.write_cell(row+in_row-1,k4+col,str(v4))
                            row += in_row
                            print(row)
        elif isinstance(response[k1],list):
            for d in response[k1]:
                print(d)