#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试GUI应用的核心功能
"""

import sys
import os
import datetime
import openpyxl

# 添加当前目录到模块搜索路径
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# 导入主程序中的函数
from main import read_product_models_from_excel, update_excel_with_test_results

def test_excel_reading():
    """测试Excel文件读取功能"""
    print("测试Excel文件读取功能...")
    
    try:
        # 读取Excel文件
        excel_filename = "density_data.xlsx"
        product_info_list = read_product_models_from_excel(excel_filename)
        
        print(f"成功读取 {len(product_info_list)} 个产品型号")
        for i, info in enumerate(product_info_list, 1):
            print(f"产品 {i}: {info}")
        
        return True
    
    except Exception as e:
        print(f"Excel文件读取失败: {str(e)}")
        return False

def test_excel_writing():
    """测试Excel文件写入功能"""
    print("\n测试Excel文件写入功能...")
    
    try:
        # 创建测试数据
        test_data = {
            "来样时间": "2023-09-18 10:00:00",
            "测试时间": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "机台号": "M001",
            "产品型号": "TestModel",
            "班次": "白班",
            "密度1": 1.2345,
            "密度2": 1.2350,
            "密度3": 1.2340,
            "密度4": 1.2342,
            "密度5": 1.2347,
            "平均值": 1.2345
        }
        
        # 更新Excel文件
        excel_filename = "density_data.xlsx"
        product_model = "TestModel"
        
        # 先将测试数据添加到Excel中
        update_excel_with_test_results(excel_filename, product_model, test_data)
        print(f"成功更新 {product_model} 的测试数据")
        
        # 验证数据是否正确写入
        wb = openpyxl.load_workbook(excel_filename)
        ws = wb.active
        
        # 查找包含TestModel的行
        updated = False
        for row in ws.iter_rows(min_row=2):
            if row[3].value == product_model:
                print(f"验证测试数据:")
                print(f"  来样时间: {row[0].value}")
                print(f"  测试时间: {row[1].value}")
                print(f"  机台号: {row[2].value}")
                print(f"  产品型号: {row[3].value}")
                print(f"  班次: {row[4].value}")
                print(f"  密度1: {row[5].value}")
                print(f"  密度2: {row[6].value}")
                print(f"  密度3: {row[7].value}")
                print(f"  密度4: {row[8].value}")
                print(f"  密度5: {row[9].value}")
                print(f"  平均值: {row[10].value}")
                updated = True
                break
        
        wb.close()
        
        return updated
    
    except Exception as e:
        print(f"Excel文件写入失败: {str(e)}")
        return False

def main():
    """主测试函数"""
    print("=== GUI应用功能测试 ===")
    
    # 测试Excel读取功能
    excel_reading_ok = test_excel_reading()
    
    # 测试Excel写入功能
    excel_writing_ok = test_excel_writing()
    
    print("\n=== 测试结果 ===")
    print(f"Excel读取功能: {'通过' if excel_reading_ok else '失败'}")
    print(f"Excel写入功能: {'通过' if excel_writing_ok else '失败'}")
    
    if excel_reading_ok and excel_writing_ok:
        print("\n✅ 所有测试通过！GUI应用的核心功能正常工作。")
        return 0
    else:
        print("\n❌ 部分测试失败！请检查应用程序。")
        return 1

if __name__ == "__main__":
    sys.exit(main())
