from openpyxl import Workbook

# 创建一个新的Excel文件
workbook = Workbook()
sheet = workbook.active

# 设置表头
headers = ["来样时间", "检测时间", "机台号", "产品型号", "班次", 
          "密度1", "密度2", "密度3", "密度4", "密度5", "平均值"]
sheet.append(headers)

# 添加检测数据
for i in range(1, 6):
    row_data = [
        f"2024-01-15 08:30:00",  # 来样时间
        "",  # 检测时间（留空，由程序填写）
        f"Machine{i:03d}",  # 机台号
        f"Model{i:03d}",  # 产品型号
        "早班",  # 班次
        "", "", "", "", "", ""  # 密度值和平均值（留空，由程序填写）
    ]
    sheet.append(row_data)

# 保存文件
workbook.save("density_data.xlsx")
print("检测Excel文件创建成功！")
