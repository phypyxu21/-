import pandas as pd
import re
"""
# 读取 Excel 文件
df = pd.read_excel('data.xlsx', engine='openpyxl')

# 显示前5行数据
print("前5行数据：")
print(df.head())

# 如果 Excel 文件有多个工作表，可以指定工作表名称或索引
# df = pd.read_excel('data.xlsx', sheet_name='Sheet1')
# 或
# df = pd.read_excel('data.xlsx', sheet_name=0)

# 显示数据的基本信息
print("\n数据基本信息：")
print(df.info())

# 显示数据的描述性统计
print("\n描述性统计：")
print(df.describe())

# 选择特定的列
ages = df['Age']
print("\n年龄列：")
print(ages)

# 筛选数据
males = df[df['Gender'] == 'Male']
print("\n男性数据：")
print(males)"""
"""pandas.read_excel(io, sheet_name=0, *, 
                  header=0, names=None, 
                  index_col=None, usecols=None, 
                  dtype=None, engine=None, 
                  converters=None, true_values=None, 
                  false_values=None, skiprows=None, 
                  nrows=None, na_values=None, 
                  keep_default_na=True, na_filter=True, 
                  verbose=False, parse_dates=False, 
                  date_parser=<no_default>, date_format=None, 
                  thousands=None, decimal='.', comment=None, 
                  skipfooter=0, storage_options=None, 
                  dtype_backend=<no_default>, engine_kwargs=None)
io：这是必需的参数，指定了要读取的 Excel 文件的路径或文件对象。

sheet_name=0：指定要读取的工作表名称或索引。默认为0，即第一个工作表。

header=0：指定用作列名的行。默认为0，即第一行。

names=None：用于指定列名的列表。如果提供，将覆盖文件中的列名。

index_col=None：指定用作行索引的列。可以是列的名称或数字。

usecols=None：指定要读取的列。可以是列名的列表或列索引的列表。

dtype=None：指定列的数据类型。可以是字典格式，键为列名，值为数据类型。

engine=None：指定解析引擎。默认为None，pandas 会自动选择。

converters=None：用于转换数据的函数字典。

true_values=None：指定应该被视为布尔值True的值。

false_values=None：指定应该被视为布尔值False的值。

skiprows=None：指定要跳过的行数或要跳过的行的列表。

nrows=None：指定要读取的行数。

na_values=None：指定应该被视为缺失值的值。

keep_default_na=True：指定是否要将默认的缺失值（例如NaN）解析为NA。

na_filter=True：指定是否要将数据转换为NA。

verbose=False：指定是否要输出详细的进度信息。

parse_dates=False：指定是否要解析日期。

date_parser=<no_default>：用于解析日期的函数。

date_format=None：指定日期的格式。

thousands=None：指定千位分隔符。

decimal='.'：指定小数点字符。

comment=None：指定注释字符。

skipfooter=0：指定要跳过的文件末尾的行数。

storage_options=None：用于云存储的参数字典。

dtype_backend=<no_default>：指定数据类型后端。

engine_kwargs=None：传递给引擎的额外参数字典。"""


"""DataFrame.to_excel(excel_writer, *, 
sheet_name='Sheet1', na_rep='',
 float_format=None, columns=None, 
 header=True, index=True, index_label=None, 
 startrow=0, startcol=0, 
 engine=None, merge_cells=True, 
 inf_rep='inf', freeze_panes=None, 
 storage_options=None, engine_kwargs=None)
 excel_writer：这是必需的参数，指定了要写入的 Excel 文件路径或文件对象。

sheet_name='Sheet1'：指定写入的工作表名称，默认为 'Sheet1'。

na_rep=''：指定在 Excel 文件中表示缺失值（NaN）的字符串，默认为空字符串。

float_format=None：指定浮点数的格式。如果为 None，则使用 Excel 的默认格式。

columns=None：指定要写入的列。如果为 None，则写入所有列。

header=True：指定是否写入列名作为第一行。如果为 False，则不写入列名。

index=True：指定是否写入索引作为第一列。如果为 False，则不写入索引。

index_label=None：指定索引列的标签。如果为 None，则不写入索引标签。

startrow=0：指定开始写入的行号，默认从第0行开始。

startcol=0：指定开始写入的列号，默认从第0列开始。

engine=None：指定写入 Excel 文件时使用的引擎，默认为 None，pandas 会自动选择。

merge_cells=True：指定是否合并单元格。如果为 True，则合并具有相同值的单元格。

inf_rep='inf'：指定在 Excel 文件中表示无穷大值的字符串，默认为 'inf'。

freeze_panes=None：指定冻结窗格的位置。如果为 None，则不冻结窗格。

storage_options=None：用于云存储的参数字典。

engine_kwargs=None：传递给引擎的额外参数字典。
"""
"""
# 1. 准备多行数据（可以是列表嵌套字典，或字典嵌套列表）
# 方式1：字典嵌套列表（键为列名，值为该列的所有数据）
data = {
    "姓名": ["张三", "李四", "王五", "赵六"],
    "年龄": [18, 17, 18, 17],
    "班级": ["高一(1)班", "高一(2)班", "高一(1)班", "高一(3)班"],
    "成绩": [95, 88, 92, 85]
}

# 2. 将数据转换为DataFrame（pandas的核心数据结构，类似表格）
df = pd.DataFrame(data)

# 3. 写入Excel文件并保存
# index=False 表示不写入行索引（否则会多一列0,1,2...）
# engine='openpyxl' 用于支持.xlsx格式
#df.to_excel("D:\zuomian\code\.vscode\python\杂七杂八的小脚本\学生信息_pandas.xlsx", index=False, engine="openpyxl")

#print("数据写入完成，文件已保存！")

df_read = pd.read_excel("D:\zuomian\code\.vscode\python\杂七杂八的小脚本\学生信息_pandas.xlsx", engine="openpyxl")
print(df_read["姓名"][0])
for i in df_read.index:
    print(f"第{i+1}行数据：")
    print(f"姓名：{df_read['姓名'][i]}，年龄：{df_read['年龄'][i]}，班级：{df_read['班级'][i]}，成绩：{df_read['成绩'][i]}")
    """
# 整合为总列表（嵌套列表）
pvc_xlpe = [
    [0.5, 0.6, "-"],
    [0.75, 0.6, 0.6],
    [1.0, 0.6, 0.6],
    [1.5, 0.7, 0.6],
    [2.5, 0.8, 0.7],
    [4, 0.8, 0.7],
    [6, 0.8, 0.7],
    [10, 1.0, 0.7]
]



GYK_read=pd.read_excel("GB-T-9330-2020kvvr工艺卡.xlsx",engine="openpyxl")
BC_NAME=GYK_read["标称截面mm2"]
bc_name_number_num_sum=[]
#提取标称截面中的数值部分
JY=GYK_read["绝缘"]
jy_B=GYK_read["标"]
jy_b=GYK_read["薄"]


JY_list=[]
jy_B_list=[]
jy_b_list=[]


for i in BC_NAME:
    bc_name_number_str_mid=re.findall(r'\d+\.?\d*',i)
    bc_name_number_num=[float(x)if'.' in x else int(x) for x in bc_name_number_str_mid]
    bc_name_number_num_sum.append(bc_name_number_num)


for i in bc_name_number_num_sum:
    for j in i:
        for k in pvc_xlpe:
            if j[1]==k[0]:
                JY_list.append(k[1])
                jy_B_list.append(k[1])
                jy_b_list.append(k[1])



