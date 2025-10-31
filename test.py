import pandas as pd
import os
import re
import math
def add_column_to_excel(file_path, sheet_name, column_name, data, overwrite=False):
    """
    向Excel文件的指定工作表和指定列添加数据
    
    参数:
    file_path: Excel文件路径
    sheet_name: 工作表名称
    column_name: 要添加数据的列名
    data: 要添加的数据列表
    overwrite: 如果列已存在，是否覆盖(True)或忽略(False)，默认False
    """
    # 检查文件是否存在
    file_exists = os.path.exists(file_path)
    
    try:
        if file_exists:
            # 读取已有Excel文件
            df = pd.read_excel(file_path, sheet_name=sheet_name, engine="openpyxl")
            
            # 检查列是否已存在
            if column_name in df.columns:
                if overwrite:
                    print(f"列 '{column_name}' 已存在，将进行覆盖")
                else:
                    print(f"列 '{column_name}' 已存在，将不会进行修改")
                    return
        else:
            # 如果文件不存在，创建一个新的DataFrame
            df = pd.DataFrame()
        
        # 确保数据长度与现有数据匹配，如果是新文件则直接创建
        if len(df) > 0 and len(data) != len(df):
            print(f"警告：数据长度({len(data)})与现有数据行数({len(df)})不匹配")
            # 可以选择截断或填充数据以匹配长度
            # 这里简单处理：如果数据短则用NaN填充，长则截断
            data = data[:len(df)] + [None]*(len(df)-len(data))
        
        # 添加数据到指定列
        df[column_name] = data
        
        # 写入Excel文件
        with pd.ExcelWriter(
            file_path,
            engine="openpyxl",
            mode="a" if file_exists else "w",
            if_sheet_exists="replace" if file_exists else None
        ) as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        print(f"成功将数据添加到列 '{column_name}'")
        
    except Exception as e:
        print(f"操作失败: {str(e)}")


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
        for k in pvc_xlpe:
            if i[1]==k[0]:
                JY_list.append(k[1])
                jy_B_list.append(k[1])
                jy_b_list.append(k[1]*0.9)

conductor_data = {
    0.5: {
        "导体名称": "5类导体",
        "截面尺寸 mm²": 0.5,
        "导体结构 mm": "28/0.15",
        "单线偏差 mm": "0.15+0.002",
        "纹线外径 mm": 0.92,
        "铜消耗 kg/km": 4.5
    },
    0.75: {
        "导体名称": "5类导体",
        "截面尺寸 mm²": 0.75,
        "导体结构 mm": "41/0.15",
        "单线偏差 mm": "0.15+0.002",
        "纹线外径 mm": 1.13,
        "铜消耗 kg/km": 6.5
    },
    1: {
        "导体名称": "5类导体",
        "截面尺寸 mm²": 1,
        "导体结构 mm": "32/0.195",
        "单线偏差 mm": "0.195±0.002",
        "纹线外径 mm": 1.29,
        "铜消耗 kg/km": 8.6
    },
    1.5: {
        "导体名称": "5类导体",
        "截面尺寸 mm²": 1.5,
        "导体结构 mm": "47/0.195",
        "单线偏差 mm": "0.195±0.002",
        "纹线外径 mm": 1.57,
        "铜消耗 kg/km": 12.6
    },
    2.5: {
        "导体名称": "5类导体",
        "截面尺寸 mm²": 2.5,
        "导体结构 mm": "76/0.195",
        "单线偏差 mm": "0.195±0.002",
        "纹线外径 mm": 2.00,
        "铜消耗 kg/km": 20.4
    },
    4: {
        "导体名称": "5类导体",
        "截面尺寸 mm²": 4,
        "导体结构 mm": "54/0.295",
        "单线偏差 mm": "0.295±0.002",
        "纹线外径 mm": 2.55,
        "铜消耗 kg/km": 33.2
    },
    6: {
        "导体名称": "5类导体",
        "截面尺寸 mm²": 6,
        "导体结构 mm": "80/0.295 8+6股×12支",
        "单线偏差 mm": "0.295±0.002",
        "纹线外径 mm": 3.42,
        "铜消耗 kg/km": 49.2
    },
    10: {
        "导体名称": "5类导体",
        "截面尺寸 mm²": 10,
        "导体结构 mm": "80/0.395 8+6股×12支",
        "单线偏差 mm": "0.395±0.002",
        "纹线外径 mm": 4.49,
        "铜消耗 kg/km": 88.2
    },
    16: {
        "导体名称": "5类导体",
        "截面尺寸 mm²": 16,
        "导体结构 mm": "126/0.40",
        "单线偏差 mm": "±0.004",
        "纹线外径 mm": 5.90,
        "铜消耗 kg/km": 145
    },
    25: {
        "导体名称": "5类导体",
        "截面尺寸 mm²": 25,
        "导体结构 mm": "196/0.40",
        "单线偏差 mm": "±0.004",
        "纹线外径 mm": 7.28,
        "铜消耗 kg/km": 226
    },
    35: {
        "导体名称": "5类导体",
        "截面尺寸 mm²": 35,
        "导体结构 mm": "276/0.40",
        "单线偏差 mm": "±0.004",
        "纹线外径 mm": 9.31,
        "铜消耗 kg/km": 318
    },
    50: {
        "导体名称": "5类导体",
        "截面尺寸 mm²": 50,
        "导体结构 mm": "396/0.40",
        "单线偏差 mm": "±0.004",
        "纹线外径 mm": 10.13,
        "铜消耗 kg/km": 456
    },
    70: {
        "导体名称": "5类导体",
        "截面尺寸 mm²": 70,
        "导体结构 mm": "360/0.50",
        "单线偏差 mm": "±0.004",
        "纹线外径 mm": 12.12,
        "铜消耗 kg/km": 648
    },
    95: {
        "导体名称": "5类导体",
        "截面尺寸 mm²": 95,
        "导体结构 mm": "475/0.50",
        "单线偏差 mm": "±0.005",
        "纹线外径 mm": 14.00,
        "铜消耗 kg/km": 855
    },
    120: {
        "导体名称": "5类导体",
        "截面尺寸 mm²": 120,
        "导体结构 mm": "608/0.50",
        "单线偏差 mm": "±0.005",
        "纹线外径 mm": 15.24,
        "铜消耗 kg/km": 1093
    },
    150: {
        "导体名称": "5类导体",
        "截面尺寸 mm²": 150,
        "导体结构 mm": "760/0.50",
        "单线偏差 mm": "±0.005",
        "纹线外径 mm": 18.25,
        "铜消耗 kg/km": 1357
    },
    185: {
        "导体名称": "5类导体",
        "截面尺寸 mm²": 185,
        "导体结构 mm": "931/0.50",
        "单线偏差 mm": "±0.005",
        "纹线外径 mm": 19.14,
        "铜消耗 kg/km": 1674
    },
    240: {
        "导体名称": "5类导体",
        "截面尺寸 mm²": 240,
        "导体结构 mm": "1216/0.50",
        "单线偏差 mm": "±0.005",
        "纹线外径 mm": 22.07,
        "铜消耗 kg/km": 2183
    }
}
#匹配绞线外径
JX_list=[]
for i in conductor_data:
    for j in bc_name_number_num_sum:
        if i==j[1]:
            JX_list.append(conductor_data[i]["纹线外径 mm"])
#控制值
Kong=[]
for i in range(len(JY_list)):
        mid=2.2*JY_list[i]+JX_list[i]
        Kong.append(mid)
#上下限
max_jy=[]
min_jy=[]
for i in range(len(Kong)):
        mid_1=Kong[i]+0.1*JY_list[i]
        mid_2=Kong[i]-0.1*JY_list[i]
        max_jy.append(mid_1)
        min_jy.append(mid_2)
#成缆系数
cabling_coefficient = {
    2: 2.00,
    3: 2.16,
    4: 2.42,
    5: 2.70,
    6: 3.00,
    7: 3.00,
    8: 3.45,
    9: 3.80,
    10: 4.00,
    11: 4.00,
    12: 4.16,
    13: 4.41,
    14: 4.41,
    15: 4.70,
    16: 4.70,
    17: 5.00,
    18: 5.00,
    19: 5.00,
    20: 5.33,
    21: 5.33,
    22: 5.67,
    23: 5.67,
    24: 6.00,
    25: 6.00,
    26: 6.00,
    27: 6.15,
    28: 5.41,
    29: 6.41,
    30: 6.41,
    31: 6.70,
    32: 6.70,
    33: 6.70,
    34: 7.00,
    35: 7.00,
    36: 7.00,
    37: 7.00,
    38: 7.33,
    39: 7.33,
    40: 7.33,
    41: 7.67,
    42: 7.67,
    43: 7.67,
    44: 8.00,
    45: 8.00,
    46: 8.00,
    47: 8.00,
    48: 8.15,
    52: 8.41,
    61: 9.00
}
#成缆系数列表
cl=[]
for i in bc_name_number_num_sum:
    for j in cabling_coefficient:
        if i[0]==j:
            cl.append(cabling_coefficient[j])
#成缆外径
cl_out=[]
for i in range(len(Kong)):
        cl_out.append(round(Kong[i]*cl[i],2))
conductor_diameter_list = [
    # 表头：[标称截面积mm², 第1种导体直径mm, 第2种导体直径mm, 第5种导体直径mm]
    [0.5, 0.8, 0.9, 1.0],
    [0.75, 1.0, 1.1, 1.1],
    [1.0, 1.1, 1.2, 1.3],
    [1.5, 1.4, 1.5, 1.5],
    [2.5, 1.8, 1.9, 2.0],
    [4, 2.2, 2.4, 2.5],
    [6, 2.7, 2.9, 3.0],
    [10, 3.5, 3.8, 3.9]
]
#计算假定直径后的护套厚度
no_real_D_list=[]
for i in bc_name_number_num_sum:
    for j in conductor_diameter_list:
        if i[1]==j[0]:
            no_real_D_list.append(j[3])
no_real_k_list=[]
for i in range(len(no_real_D_list)):
        mid=no_real_D_list[i]+2*JY_list[i]
        no_real_k_list.append(mid)
no_real_cl_out=[]
for i in range(len(no_real_k_list)):
        no_real_mid=round(no_real_k_list[i]*cl[i]+0.2,2)
        no_real_cl_out.append(no_real_mid)
HT_list=[]
for i in no_real_cl_out:
    if i<=10:
        HT_list.append(1.2)
    elif i>10 and i<=16:
        HT_list.append(1.5)
    elif i>16 and i<=25:
        HT_list.append(1.7)
    elif i>25 and i<=30:
        HT_list.append(2.0)
    elif i>30 and i<=40:
        HT_list.append(2.2)
    elif i>40:
        HT_list.append(2.5)
#绝缘火花
jy_fire=[]
for i in JY_list:
        if i<=0.25:
            
            jy_fire.append("3kV")
        elif 0.25<i<=0.5:
            
            jy_fire.append("4kV")
        elif 0.5<i<=1:
            
            jy_fire.append("6kV")
        elif 1<i<=1.5:
            
            jy_fire.append("10kV")
        elif 1.5<i<=2.0:
            
            jy_fire.append("15kV")
        elif 2.0<i<=2.5:
            
            jy_fire.append("20kV")
        elif 2.5<i:
            
            jy_fire.append("25kV")
ht_fire=[]
for i in HT_list:
    mid=i*6
    if mid>15:
        ht_fire.append("15kV")
    else:
        ht_fire.append(mid)
#绝缘模芯直径
MX=[]
for i in JX_list:
    mid = i+0.2
    MX.append(round(mid,1))
#绝缘模套直径等于控制值

#护套模芯
HS_MX=[]
for i in cl_out:
    mid=i+4#i+3~6都可以
    HS_MX.append(round(mid,1))
#护套模套
HS_MT=[]
for i in range(len(HS_MX)):
    mid=HS_MX[i]+1+3*HT_list[i]
    HS_MT.append(round(mid,1))

#铜消耗
coper_use=[]
for i in bc_name_number_num_sum:
    for j in conductor_data:
        if i[1]==j:
            coper_use.append(conductor_data[j]["铜消耗 kg/km"]*i[0])
#绝缘消耗
insulation_use=[]
for i in range(len(bc_name_number_num_sum)):
        #绝缘有平均值要求要×1.05
        mid=(1.05*JY_list[i]+JX_list[i])*JY_list[i]*1.4*3.14156*1.021*1.05*bc_name_number_num_sum[i][0]
        insulation_use.append(round(mid,2))
#护套消耗
ht_use=[]
for i in range(len(bc_name_number_num_sum)):
        mid=(cl_out[i]+HT_list[i])*HT_list[i]*1.4*3.14156
        ht_use.append(round(mid,2))
# 示例用法
if __name__ == "__main__":
    # 要添加的数据
    numbers = ht_use
    
    # 调用函数添加数据
    add_column_to_excel(
        file_path="D:\Desktop\杂七杂八的小脚本\-\GB-T-9330-2020kvvr工艺卡.xlsx",    # Excel文件路径
        sheet_name="Sheet1",      # 工作表名称
        column_name="护套",     # 要添加到的列名
        data=numbers,             # 要添加的数据
        overwrite=True           # 如果列存在，覆盖
    )
