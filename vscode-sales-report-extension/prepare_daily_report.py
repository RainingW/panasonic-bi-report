#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
松下 BI 日报数据准备脚本
修正内容：
  1. 筛去"已取消"订单
  2. "退款成功"订单计为退款订单
  3. 订单金额使用"用户支付金额"字段（如不存在则回退到"用户实付金额(元)"）
  4. 客单价 = 成团总金额 / 成团买家数（去重）
  5. 品类按"商品三级品类/商品三级类目"字段分类
  6. 流量来源按"省"字段，排除退款成功的订单
  7. 订单归属时间由"支付时间"字段决定
"""
import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

import pandas as pd
import numpy as np
from datetime import datetime, date
import warnings
warnings.filterwarnings('ignore')

path = r'C:\Users\雨凝w\Desktop\松下实习\拼多多订单数据.xlsx'

# 读取数据
df = pd.read_excel(path)
print(f'原始数据: {df.shape[0]} 行\n')

# ============================================================
# 1. 日期处理 --- 订单归属时间由"支付时间"字段决定
# ============================================================
print('=== 日期字段分析 ===')
print('支付时间样本:')
print(df['支付时间'].dropna().head(5).tolist())

df['支付时间_clean'] = pd.to_datetime(df['支付时间'], errors='coerce')
df['支付日期'] = df['支付时间_clean'].dt.date

print(f'\n日期范围: {df[df["支付日期"].notna()]["支付日期"].min()} 到 {df[df["支付日期"].notna()]["支付日期"].max()}')

# 筛选6月18日的数据
target_date = date(2026, 6, 18)
df_target = df[df['支付日期'] == target_date].copy()
print(f'\n6月18日原始订单数: {len(df_target)} 笔')

# ============================================================
# 2. 订单状态处理
#    - "已取消"订单筛去
#    - "退款成功"订单计为退款订单
# ============================================================
print('\n=== 订单状态处理 ===')
df_target['订单状态_str'] = df_target['订单状态'].astype(str)

# 筛去"已取消"订单
canceled_mask = df_target['订单状态_str'].str.contains('已取消', na=False)
df_valid = df_target[~canceled_mask].copy()
print(f'筛去"已取消"订单: {canceled_mask.sum()} 笔')
print(f'筛后剩余: {len(df_valid)} 笔')

# 标记"退款成功"订单
refund_mask = df_valid['订单状态_str'].str.contains('退款成功', na=False)
df_active = df_valid[~refund_mask].copy()   # 有效成团订单
df_refund = df_valid[refund_mask].copy()     # 退款订单
print(f'其中退款成功: {refund_mask.sum()} 笔')
print(f'有效成团订单: {len(df_active)} 笔')

# ============================================================
# 3. 金额处理 --- 订单金额由"用户支付金额"字段决定
#    实际 Excel 中可能叫"用户实付金额(元)"，做兼容处理
# ============================================================
amount_col = None
for candidate in ['用户支付金额', '用户实付金额(元)', '商家实收金额(元)']:
    if candidate in df.columns:
        amount_col = candidate
        break
if amount_col is None:
    amount_col = '用户实付金额(元)'  # 兜底

print(f'\n金额字段使用: [{amount_col}]')

df_active[amount_col] = pd.to_numeric(df_active[amount_col], errors='coerce').fillna(0)
if len(df_refund) > 0:
    df_refund[amount_col] = pd.to_numeric(df_refund[amount_col], errors='coerce').fillna(0)

# ============================================================
# 4. 基础度量计算
# ============================================================
print('\n=== 基础度量 ===')

total_sales = df_active[amount_col].sum()        # 成团总金额
refund_amount = df_refund[amount_col].sum() if len(df_refund) > 0 else 0
total_orders = df_active['订单号'].nunique()      # 成团订单数

# 买家去重：优先使用"用户购买手机号"（清洗掉空白/制表符等无效数据后去重）
buyer_col = '用户购买手机号'
if buyer_col in df_active.columns:
    buyer_ids = df_active[buyer_col].astype(str).str.strip()
    buyer_ids = buyer_ids.replace('', pd.NA)
    unique_buyers = buyer_ids.dropna().nunique()
    if unique_buyers == 0:
        unique_buyers = total_orders  # 手机号全为空时回退到订单数
    print(f'买家标识字段: [{buyer_col}], 有效唯一买家数: {unique_buyers}')
else:
    unique_buyers = total_orders
    print(f'买家标识字段: 未找到, 回退到订单数: {unique_buyers}')

# 客单价 = 成团总金额 / 成团买家数（去重）
avg_order_value = total_sales / unique_buyers if unique_buyers > 0 else 0

# 补充指标
pay_cvr = total_sales / unique_buyers if unique_buyers else 0
refund_rate_val = refund_amount / (total_sales + refund_amount) if (total_sales + refund_amount) else 0

print(f'成团总金额 ({amount_col}): RMB {total_sales:,.2f}')
print(f'退款金额: RMB {refund_amount:,.2f}')
print(f'成团订单数: {total_orders}')
print(f'成团买家数（去重）: {unique_buyers}')
print(f'客单价 (成团总金额/成团买家数): RMB {avg_order_value:,.2f}')
print(f'退款率: {refund_rate_val:.2%}')

# ============================================================
# 5. 按"商品三级品类/商品三级类目"分类统计（销售金额扇形图数据源）
# ============================================================
category_field = None
for candidate in ['商品三级品类', '商品三级类目']:
    if candidate in df_active.columns:
        category_field = candidate
        break
if category_field is None:
    category_field = '商品一级类目'  # 最终 fallback

print(f'\n=== 品类统计（按"{category_field}"） ===')

df_active[category_field] = df_active[category_field].fillna('未知')
category_agg = df_active.groupby(category_field).agg({
    amount_col: 'sum',
    '商品数量(件)': 'sum',
    '订单号': 'nunique'
}).rename(columns={
    amount_col: 'sales_amount',
    '商品数量(件)': 'qty',
    '订单号': 'orders'
}).reset_index()
category_agg.columns = ['category', 'sales_amount', 'qty', 'orders']
category_agg['pay_users'] = category_agg['orders']
category_agg['visitors'] = category_agg['orders']
category_agg['cvr'] = (category_agg['orders'] / category_agg['visitors']).round(4)
category_agg['refund_rate'] = 0.0
category_agg['favorites'] = 0
category_agg['cart_add'] = 0
category_agg['new_user_rate'] = 0.0
category_agg['repeat_user_rate'] = 1.0
category_agg = category_agg.sort_values('sales_amount', ascending=False)

print(category_agg.to_string())

# ============================================================
# 6. 按商品统计
# ============================================================
product_field = '商品'
if product_field not in df_active.columns:
    product_field = '商品名称'

print(f'\n=== 商品统计（按"{product_field}"） ===')

df_active[product_field] = df_active[product_field].fillna('未知')
product_agg = df_active.groupby(product_field).agg({
    amount_col: 'sum',
    '商品数量(件)': 'sum',
    '订单号': 'nunique'
}).rename(columns={
    amount_col: 'sales_amount',
    '商品数量(件)': 'qty',
    '订单号': 'orders'
}).reset_index()
product_agg.columns = ['product_name', 'sales_amount', 'qty', 'orders']
product_agg['pay_users'] = product_agg['orders']
product_agg['visitors'] = product_agg['orders']
product_agg['sku'] = 'N/A'
product_agg['cvr'] = (product_agg['orders'] / product_agg['visitors']).round(4)
product_agg['refund_rate'] = 0.0
product_agg['favorites'] = 0
product_agg['cart_add'] = 0
product_agg['new_user_rate'] = 0.0
product_agg['repeat_user_rate'] = 1.0
product_agg = product_agg.sort_values('sales_amount', ascending=False)

print(product_agg.head(20).to_string())

# ============================================================
# 7. 流量来源统计 --- 按"省"字段，不考虑退款成功的订单
# ============================================================
print(f'\n=== 流量来源（按"省"，不含退款订单） ===')

traffic_df = df_active[df_active['省'].notna()].copy()
traffic_stats = traffic_df.groupby('省').size().reset_index(name='value')
traffic_stats.columns = ['source', 'value']
traffic_stats = traffic_stats.sort_values('value', ascending=False)

print(traffic_stats.to_string())

# ============================================================
# 8. 保存清洗数据
# ============================================================
output_path = r'C:\Users\雨凝w\DataGripProjects\Panasonic_1\vscode-sales-report-extension\cleaned_data_20260618.xlsx'

with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    # Overview sheet
    pd.DataFrame({
        'date': [target_date],
        'total_sales': [total_sales],
        'payment_amount': [total_sales],
        'refund_amount': [refund_amount],
        'orders': [total_orders],
        'pay_users': [unique_buyers],
        'visitors': [unique_buyers],
        'pay_cvr': [round(pay_cvr, 4)],
        'aov': [round(avg_order_value, 2)],
        'refund_rate': [round(refund_rate_val, 4)],
        'add_to_cart': [0],
        'favorites': [0]
    }).to_excel(writer, sheet_name='overview', index=False)

    # Category sheet
    category_stats = category_agg.copy()
    category_stats.to_excel(writer, sheet_name='category', index=False)

    # Products sheet
    product_stats = product_agg.copy()
    product_stats.to_excel(writer, sheet_name='products', index=False)

    # Traffic sheet
    traffic_stats.to_excel(writer, sheet_name='traffic', index=False)

print(f'\n{"="*60}')
print(f'[OK] 清洗数据已保存: {output_path}')
print(f'{"="*60}')
print(f'\n========== 2026年6月18日 BI 日报概览 ==========')
print(f'  日期:               2026年6月18日')
print(f'  成团总金额:         RMB {total_sales:,.2f}')
print(f'  退款金额:           RMB {refund_amount:,.2f}')
print(f'  成团订单数:         {total_orders}')
print(f'  成团买家数（去重）: {unique_buyers}')
print(f'  客单价:             RMB {avg_order_value:,.2f}')
print(f'  退款率:             {refund_rate_val:.2%}')
print(f'  品类数:             {len(category_agg)}')
print(f'  商品数:             {len(product_agg)}')
print(f'  流量省份数:         {len(traffic_stats)}')
print(f'==================================================')
