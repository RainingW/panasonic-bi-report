#!/usr/bin/env python
# -*- coding: utf-8 -*-
import pandas as pd
import sys

path = r'C:\Users\雨凝w\Desktop\松下实习\拼多多订单数据.xlsx'
try:
    df = pd.read_excel(path)
    print('✓ 文件加载成功')
    print(f'数据形状: {df.shape[0]} 行 × {df.shape[1]} 列\n')
    print('列名:')
    for i, col in enumerate(df.columns):
        print(f'  {i+1}. {col}')
    print('\n前5行数据:')
    print(df.head(5).to_string())
    print('\n\n数据类型:')
    print(df.dtypes)
    print('\n\n缺失值统计:')
    print(df.isnull().sum())
    print('\n\n样本统计:')
    print(df.describe())
except Exception as e:
    print(f'✗ 错误: {e}', file=sys.stderr)
    sys.exit(1)
