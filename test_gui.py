#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试GUI功能的简单脚本
"""

import os
import pandas as pd

def create_test_files():
    """创建测试文件"""
    # 创建data目录
    os.makedirs('./data', exist_ok=True)
    os.makedirs('./test', exist_ok=True)
    
    # 创建测试数据
    data1 = {
        'A': [1, 2, 3, 4, 5],
        'B': ['电压', '电流', '功率', '频率', '温度'],
        'C': [11, 12, 13, 14, 15],
        'F': [23, 24, 34, 4, 54]
    }
    
    data2 = {
        'A': [1, 2, 5, 4, 5, 3],
        'B': ['电压', '电流', '温度', '频率', '温度', '电流C'],
        'C': [16, 17, 18, 19, 20, 8],
        'E': [4, 56, 6, 23, 23, 12]
    }
    
    # 创建DataFrame并保存
    df1 = pd.DataFrame(data1)
    df2 = pd.DataFrame(data2)
    
    # 保存为Excel和CSV
    df1.to_excel('./test/1.xlsx', index=False)
    df2.to_csv('./test/2.csv', index=False, encoding='utf-8-sig')
    
    print("测试文件已创建:")
    print("- ./test/1.xlsx")
    print("- ./test/2.csv")
    print("- ./data/ 目录已准备就绪")

if __name__ == "__main__":
    create_test_files()
