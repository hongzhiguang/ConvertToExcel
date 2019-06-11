import  os

# 工程目录
proj_path = os.path.dirname(os.path.dirname(__file__))

# 提取数据的原文件
data_path = os.path.normpath(os.path.join(proj_path,"data","data.py"))

# 提取数据后存入的excel表格
excel_path = os.path.normpath(os.path.join(proj_path,"data","resData.xlsx"))
