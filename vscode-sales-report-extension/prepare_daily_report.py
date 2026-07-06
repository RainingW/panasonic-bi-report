#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
松下 BI 日报数据准备脚本 v2
集成双数据源:
  1. 拼多多订单数据.xlsx — 订单级明细（成团金额/订单数/品类/商品/省份，以该数据为准）
  2. 松下冰箱pdd京东后台数据.xlsx — 后台每日汇总（访客/下单/支付/成团漏斗 + 趋势）
"""
import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

import pandas as pd
import numpy as np
from datetime import datetime, date, timedelta
import warnings
warnings.filterwarnings('ignore')

# ============================================================
# 配置
# ============================================================
ORDER_DATA_PATH = r'C:\Users\雨凝w\Desktop\松下实习\拼多多订单数据.xlsx'
BACKEND_DATA_PATH = r'C:\Users\雨凝w\Desktop\松下实习\松下冰箱pdd京东后台数据.xlsx'
OUTPUT_DIR = r'C:\Users\雨凝w\DataGripProjects\Panasonic_1\vscode-sales-report-extension'

target_date = date(2026, 6, 18)
output_path = OUTPUT_DIR + r'\cleaned_data_20260618.xlsx'

def excel_serial_to_date(serial):
    """将 Excel 序列号转为 date"""
    try:
        return (datetime(1899, 12, 30) + timedelta(days=int(serial))).date()
    except:
        return None

# ============================================================
# 第一部分：读取订单数据（作为金额/订单统计的主数据源）
# ============================================================
print('=' * 60)
print('第一部分：订单明细数据处理（拼多多订单数据.xlsx）')
print('=' * 60)

df_order = pd.read_excel(ORDER_DATA_PATH)
print(f'原始订单: {df_order.shape[0]} 行 × {df_order.shape[1]} 列')

# 日期处理
df_order['支付时间_clean'] = pd.to_datetime(df_order['支付时间'], errors='coerce')
df_order['支付日期'] = df_order['支付时间_clean'].dt.date
print(f'日期范围: {df_order["支付日期"].dropna().min()} ~ {df_order["支付日期"].dropna().max()}')

# 筛选目标日期
df_target = df_order[df_order['支付日期'] == target_date].copy()
print(f'\n{target_date} 原始订单: {len(df_target)} 笔')

if len(df_target) == 0:
    print(f'[ERROR] {target_date} 无订单数据！')
    sys.exit(1)

# 订单状态分类
df_target['订单状态_str'] = df_target['订单状态'].astype(str)
df_target['is_canceled'] = df_target['订单状态_str'].str.contains('已取消', na=False)
df_target['is_refunded'] = df_target['订单状态_str'].str.contains('退款成功', na=False)
df_target['is_active'] = ~df_target['is_canceled'] & ~df_target['is_refunded']

canceled_count = df_target['is_canceled'].sum()
refund_count = df_target['is_refunded'].sum()

# 注意：拼多多数据中"退款成功"的订单可能同时标记为"已取消"——
# 已取消的排在最前面筛掉，退款成功的从剩余中识别
# 所以需要先筛已取消，再从剩余中识别退款
df_valid = df_target[~df_target['is_canceled']].copy()  # 排除已取消
df_active = df_valid[~df_valid['is_refunded']].copy()    # 有效成团
df_refund = df_valid[df_valid['is_refunded']].copy()     # 退款订单

print(f'  已取消: {canceled_count} 笔')
print(f'  退款成功: {len(df_refund)} 笔')
print(f'  有效成团: {len(df_active)} 笔')

# 金额字段
amount_col = None
for c in ['用户支付金额', '用户实付金额(元)', '商家实收金额(元)']:
    if c in df_order.columns:
        amount_col = c
        break
if amount_col is None:
    amount_col = '用户实付金额(元)'
print(f'金额字段: [{amount_col}]')

df_active[amount_col] = pd.to_numeric(df_active[amount_col], errors='coerce').fillna(0)
if len(df_refund) > 0:
    df_refund[amount_col] = pd.to_numeric(df_refund[amount_col], errors='coerce').fillna(0)

# === 基础度量 ===
total_sales = df_active[amount_col].sum()
refund_amount = df_refund[amount_col].sum() if len(df_refund) > 0 else 0
total_orders = df_active['订单号'].nunique()

# 买家去重（拼多多手机号字段全为 \t，不可用 → 回退到订单数）
buyer_col = '用户购买手机号'
if buyer_col in df_active.columns:
    buyer_ids_clean = df_active[buyer_col].astype(str).str.strip().replace('', pd.NA)
    unique_buyers = buyer_ids_clean.dropna().nunique()
else:
    unique_buyers = 0
if unique_buyers == 0:
    unique_buyers = total_orders

aov = total_sales / unique_buyers if unique_buyers else 0
refund_rate = refund_amount / (total_sales + refund_amount) if (total_sales + refund_amount) else 0

print(f'\n--- 基础度量（以订单数据为准）---')
print(f'成团金额: ¥{total_sales:,.2f}')
print(f'退款金额: ¥{refund_amount:,.2f}')
print(f'成团订单数: {total_orders}')
print(f'成团买家数(订单数兜底): {unique_buyers}')
print(f'客单价: ¥{aov:,.2f}')
print(f'退款率: {refund_rate:.2%}')

# ============================================================
# 第二部分：读取后台数据（漏斗 + 趋势）
# ============================================================
print(f'\n{"=" * 60}')
print('第二部分：后台数据处理（松下冰箱pdd京东后台数据.xlsx）')
print('=' * 60)

df_backend = pd.read_excel(BACKEND_DATA_PATH, sheet_name='拼多多')
df_backend['日期_dt'] = df_backend['日期'].apply(excel_serial_to_date)

# 查找目标日期行
backend_row = df_backend[df_backend['日期_dt'] == target_date]
has_backend = len(backend_row) > 0

if has_backend:
    br = backend_row.iloc[0]
    print(f'找到 {target_date} 后台数据')
    funnel_visitors = int(br['商品访客数'])
    funnel_orders = int(br['下单人数'])
    funnel_payers = int(br['支付人数'])
    funnel_success = int(br['成团人数'])
    funnel_inquiry = int(br['询单客户数']) if pd.notna(br['询单客户数']) else 0

    # 漏斗转化率（来自后台）
    order_rate = float(br['下单率']) if pd.notna(br['下单率']) else 0
    pay_rate = float(br['支付率']) if pd.notna(br['支付率']) else 0
    success_rate = float(br['成团率']) if pd.notna(br['成团率']) else 0
    full_cvr = float(br['全店转化率']) if pd.notna(br['全店转化率']) else 0
    inquiry_rate = float(br['询单转化率']) if pd.notna(br['询单转化率']) else 0
    backend_refund_rate = float(br['退款率']) if pd.notna(br['退款率']) else 0

    # 后台的金额（用于对比参考）
    backend_sales = float(br['成团金额（元）'])
    backend_refund_amt = float(br['退款金额'])
    backend_net = float(br['去退金额'])
    backend_order_count = int(br['成团订单数'])
    backend_refund_count = int(br['退款单数'])

    print(f'  后台成团金额: ¥{backend_sales:,.2f} (参考)')
    print(f'  后台成团单数: {backend_order_count} (含退款订单)')
    print(f'  订单数据成团金额: ¥{total_sales:,.2f} (采用)')
    print(f'  订单数据成团单数: {total_orders} (已剔除退款)')

    print(f'\n  漏斗: 访客{funnel_visitors} → 下单{funnel_orders} → 支付{funnel_payers} → 成团{funnel_success}')
    print(f'  下单率: {order_rate:.2%}  支付率: {pay_rate:.2%}  成团率: {success_rate:.2%}  全店转化率: {full_cvr:.2%}')
else:
    print(f'[WARNING] 后台数据中未找到 {target_date}！漏斗将为空。')
    funnel_visitors = funnel_orders = funnel_payers = funnel_success = funnel_inquiry = 0
    order_rate = pay_rate = success_rate = full_cvr = inquiry_rate = backend_refund_rate = 0
    backend_sales = backend_refund_amt = backend_net = 0
    backend_order_count = backend_refund_count = 0

# 近30天趋势（从后台数据）
trend_start = target_date - timedelta(days=29)
trend_data = df_backend[(df_backend['日期_dt'] >= trend_start) & (df_backend['日期_dt'] <= target_date)].copy()
trend_data = trend_data.sort_values('日期_dt')

trend_days = [str(d) for d in trend_data['日期_dt'].tolist()]
trend_visitors = [int(v) if pd.notna(v) else 0 for v in trend_data['商品访客数'].tolist()]
trend_sales = [float(s) if pd.notna(s) else 0 for s in trend_data['成团金额（元）'].tolist()]
trend_orders_list = [int(o) if pd.notna(o) else 0 for o in trend_data['成团订单数'].tolist()]
trend_cvr_list = [float(c) if pd.notna(c) else 0 for c in trend_data['全店转化率'].tolist()]

print(f'\n近30天趋势: {len(trend_data)} 天数据 ({trend_start} ~ {target_date})')

# ============================================================
# 第三部分：品类统计（订单数据）
# ============================================================
print(f'\n{"=" * 60}')
print('第三部分：品类统计')
print('=' * 60)

cat_field = None
for c in ['商品三级类目', '商品三级品类', '商品二级类目', '商品一级类目']:
    if c in df_active.columns:
        cat_field = c
        break
if not cat_field:
    cat_field = '商品一级类目'

df_active[cat_field] = df_active[cat_field].fillna('未知')

# 品类退款统计 (from refund orders)
refund_by_cat = pd.Series(dtype=float)
if len(df_refund) > 0:
    df_refund[cat_field] = df_refund[cat_field].fillna('未知')
    refund_by_cat = df_refund.groupby(cat_field)[amount_col].sum()

category_agg = df_active.groupby(cat_field).agg(
    sales_amount=(amount_col, 'sum'),
    qty=('商品数量(件)', 'sum'),
    orders=('订单号', 'nunique')
).reset_index()
category_agg.columns = ['category', 'sales_amount', 'qty', 'orders']

# 品类退款金额
category_agg['refund_amount'] = category_agg['category'].map(refund_by_cat).fillna(0)
category_agg['refund_rate'] = (category_agg['refund_amount'] / (category_agg['sales_amount'] + category_agg['refund_amount'])).round(4)
category_agg['sales_share'] = (category_agg['sales_amount'] / total_sales).round(4) if total_sales else 0
category_agg['pay_users'] = category_agg['orders']
category_agg = category_agg.sort_values('sales_amount', ascending=False)

for _, r in category_agg.iterrows():
    print(f'  {r["category"]}: ¥{r["sales_amount"]:,.0f} ({r["orders"]}单, 退款率{r["refund_rate"]:.1%})')

# ============================================================
# 第四部分：商品统计（订单数据）
# ============================================================
print(f'\n{"=" * 60}')
print('第四部分：商品统计')
print('=' * 60)

prod_field = '商品'
if prod_field not in df_active.columns:
    prod_field = '商品名称'

df_active[prod_field] = df_active[prod_field].fillna('未知')

# 商品退款统计
refund_by_prod = pd.Series(dtype=float)
if len(df_refund) > 0:
    df_refund[prod_field] = df_refund[prod_field].fillna('未知')
    refund_by_prod = df_refund.groupby(df_refund[prod_field])[amount_col].sum()

product_agg = df_active.groupby(prod_field).agg(
    sales_amount=(amount_col, 'sum'),
    qty=('商品数量(件)', 'sum'),
    orders=('订单号', 'nunique')
).reset_index()
product_agg.columns = ['product_name', 'sales_amount', 'qty', 'orders']

# SKU
if '商品id' in df_active.columns:
    sku_map = df_active.groupby(prod_field)['商品id'].first()
    product_agg['sku'] = product_agg['product_name'].map(sku_map).fillna('N/A').astype(str)
else:
    product_agg['sku'] = 'N/A'

product_agg['pay_users'] = product_agg['orders']
product_agg['avg_price'] = (product_agg['sales_amount'] / product_agg['qty'].clip(lower=1)).astype(int)
product_agg['refund_amount'] = product_agg['product_name'].map(refund_by_prod).fillna(0)
product_agg['refund_rate'] = (product_agg['refund_amount'] / (product_agg['sales_amount'] + product_agg['refund_amount'])).round(4)
product_agg = product_agg.sort_values('sales_amount', ascending=False)

print(f'商品数: {len(product_agg)}')
for _, r in product_agg.head(10).iterrows():
    print(f'  {r["product_name"][:40]}: ¥{r["sales_amount"]:,.0f} ({r["qty"]}件) 退款率{r["refund_rate"]:.1%}')

# ============================================================
# 第五部分：流量来源（订单数据，按省份）
# ============================================================
print(f'\n{"=" * 60}')
print('第五部分：流量来源（按省份）')
print('=' * 60)

traffic_stats = pd.DataFrame()
if '省' in df_active.columns:
    traffic_df = df_active[df_active['省'].notna()].copy()
    traffic_df['省'] = traffic_df['省'].astype(str).str.strip()
    traffic_df.loc[traffic_df['省'] == '****', '省'] = '其他(数据脱敏)'
    traffic_stats = traffic_df.groupby('省').agg(
        value=('订单号', 'nunique'),
        sales=('用户实付金额(元)', 'sum')
    ).reset_index()
    traffic_stats.columns = ['source', 'value', 'sales']
    traffic_stats = traffic_stats.sort_values('value', ascending=False)

print(traffic_stats.to_string() if len(traffic_stats) > 0 else '无省份数据')

# ============================================================
# 第六部分：生成专业洞察
# ============================================================
print(f'\n{"=" * 60}')
print('第六部分：生成专业洞察')
print('=' * 60)

insights = []

# 1. 漏斗健康度分析
if has_backend:
    if full_cvr < 0.01:
        insights.append({
            'level': 'red', 'title': '全店转化率严重偏低',
            'text': f'全店转化率仅 {full_cvr:.2%}（访客→成团），低于家电行业均值3-5%。核心瓶颈在"访客→下单"环节（下单率仅{order_rate:.2%}），需重点优化主图、详情页与价格竞争力。'
        })
    elif full_cvr < 0.02:
        insights.append({
            'level': 'yellow', 'title': '全店转化率有提升空间',
            'text': f'全店转化率 {full_cvr:.2%}（访客→成团），距行业均值仍有差距。下单率 {order_rate:.2%}，建议通过评价优化、限时促销、主图A/B测试提升转化。'
        })
    else:
        insights.append({
            'level': 'green', 'title': '全店转化率良好',
            'text': f'全店转化率 {full_cvr:.2%}，访客→成团链路效率处于健康水平。继续保持内容优化与精准投放。'
        })

    if pay_rate < 0.6:
        insights.append({
            'level': 'red', 'title': '下单→支付流失严重',
            'text': f'下单 {funnel_orders} 人 → 支付 {funnel_payers} 人，支付率仅 {pay_rate:.2%}。{funnel_orders - funnel_payers} 人下单未付，建议配置购物车催付提醒（短信/推送）、限时优惠倒计时。'
        })
    elif pay_rate < 0.8:
        insights.append({
            'level': 'yellow', 'title': '下单→支付有优化空间',
            'text': f'支付率 {pay_rate:.2%}，仍有 {funnel_orders - funnel_payers} 人下单后放弃。建议在结算页增加信任标识（7天无理由/正品保证），减少支付摩擦。'
        })
    else:
        insights.append({
            'level': 'green', 'title': '下单→支付转化健康',
            'text': f'支付率 {pay_rate:.2%}，下单用户中大部分完成支付，结算体验良好。'
        })

    if funnel_inquiry > 0:
        insights.append({
            'level': 'blue', 'title': '客服询单转化分析',
            'text': f'当日询单 {funnel_inquiry} 人，询单转化率 {inquiry_rate:.1%}。高客单价家电品类客服介入对转化至关重要，建议加强客服专业培训与响应速度。'
        })

# 2. 退款分析
if refund_rate > 0.25:
    insights.append({
        'level': 'red', 'title': '退款率严重偏高',
        'text': f'退款率高达 {refund_rate:.1%}（退款¥{refund_amount:,.0f}/成团¥{total_sales:,.0f}），远超家电行业正常水平(1-5%)。需紧急排查：①商品质量/描述不符 ②物流破损 ③售后处理周期。建议逐单分析退款原因并建立预警机制。'
    })
elif refund_rate > 0.10:
    insights.append({
        'level': 'red', 'title': '退款率偏高',
        'text': f'退款率 {refund_rate:.1%}，需引起重视。建议分类分析退款原因（质量/物流/价格/预期不符），针对TOP退款商品优先改进。'
    })
elif refund_rate > 0.05:
    insights.append({
        'level': 'yellow', 'title': '退款率略高',
        'text': f'退款率 {refund_rate:.1%}，略高于健康水平。持续关注退款原因分布，重点监控高退款SKU。'
    })
else:
    insights.append({
        'level': 'green', 'title': '退款率正常',
        'text': f'退款率 {refund_rate:.1%}，处于健康水平。'
    })

# 3. 客单价分析
if aov > 3500:
    insights.append({
        'level': 'blue', 'title': '高客单价优势明显',
        'text': f'客单价 ¥{aov:,.0f}，属于高客单家电品类。建议：①推出分期免息服务降低决策门槛 ②强化延保/安装等增值服务提升ARPU ③用内容营销（评测/安装案例）增强信任。'
    })
elif aov > 2000:
    insights.append({
        'level': 'blue', 'title': '客单价中等偏上',
        'text': f'客单价 ¥{aov:,.0f}，符合冰箱品类特征。可考虑搭配周边产品（除味剂/保鲜盒）提升连带率。'
    })
else:
    insights.append({
        'level': 'yellow', 'title': '客单价偏低',
        'text': f'客单价 ¥{aov:,.0f}。对于松下品牌冰箱品类偏低，检查是否促销折扣过大或低端SKU占比过高。'
    })

# 4. 品类结构分析
if len(category_agg) > 0:
    top_cat = category_agg.iloc[0]
    cat_concentration = top_cat['sales_amount'] / total_sales if total_sales else 0
    if cat_concentration > 0.6:
        insights.append({
            'level': 'yellow', 'title': '品类集中度偏高',
            'text': f'"{top_cat["category"]}"占销售额{cat_concentration:.0%}，依赖度过高存在风险。建议培育第二增长品类，分散单品/品类波动风险。'
        })
    else:
        insights.append({
            'level': 'green', 'title': '品类结构均衡',
            'text': f'品类分布相对均衡，TOP1品类占比{cat_concentration:.0%}。可针对各品类制定差异化运营策略。'
        })

    # 高退款品类
    high_refund_cat = category_agg[category_agg['refund_rate'] > 0.2]
    if len(high_refund_cat) > 0:
        names = '、'.join(high_refund_cat['category'].head(3).tolist())
        insights.append({
            'level': 'red', 'title': '高风险品类预警',
            'text': f'"{names}"退款率偏高，建议重点排查该品类商品质量、详情页描述准确性和物流包装。'
        })

# 5. 流量效率
if has_backend and funnel_visitors > 0:
    visitor_value = total_sales / funnel_visitors
    insights.append({
        'level': 'blue', 'title': '流量效率分析',
        'text': f'单访客价值 ¥{visitor_value:.2f}（成团金额/访客数）。参考标准：冰箱品类单访客价值通常在¥3-8。当前{"高于" if visitor_value > 5 else "低于"}均值，建议{"维持高意向流量获取策略" if visitor_value > 5 else "优化流量精准度和详情页转化能力"}。'
    })

# 6. 趋势判断
if len(trend_data) >= 7:
    recent_7 = trend_data.tail(7)
    prev_7 = trend_data.iloc[-14:-7] if len(trend_data) >= 14 else trend_data.iloc[:-7]
    if len(prev_7) > 0:
        recent_avg_sales = recent_7['成团金额（元）'].mean()
        prev_avg_sales = prev_7['成团金额（元）'].mean()
        if prev_avg_sales > 0:
            change = (recent_avg_sales - prev_avg_sales) / prev_avg_sales
            if change > 0.15:
                insights.append({
                    'level': 'green', 'title': '近7日销售额上升趋势',
                    'text': f'近7日均销 ¥{recent_avg_sales:,.0f}，较前7日增长{change:.0%}。趋势向好，建议维持当前运营节奏并观察可持续性。'
                })
            elif change < -0.10:
                insights.append({
                    'level': 'red', 'title': '近7日销售额下滑',
                    'text': f'近7日均销 ¥{recent_avg_sales:,.0f}，较前7日下降{abs(change):.0%}。需排查：①竞品活动 ②流量下降 ③转化率恶化 ④618后疲软期。'
                })
            else:
                insights.append({
                    'level': 'blue', 'title': '近7日销售平稳',
                    'text': f'近7日均销 ¥{recent_avg_sales:,.0f}，变化{change:+.0%}，销售趋势稳定。'
                })

# 7. TOP商品分析
if len(product_agg) > 0:
    top_prod = product_agg.iloc[0]
    insights.append({
        'level': 'blue', 'title': f'TOP1商品: {top_prod["product_name"][:25]}',
        'text': f'销售额 ¥{top_prod["sales_amount"]:,.0f}（{top_prod["qty"]}件），占总销售{top_prod["sales_amount"]/total_sales:.0%}。{"退款率偏高，需关注售后" if top_prod["refund_rate"] > 0.1 else "表现稳健，建议保障库存与曝光"}。'
    })

# 8. 综合行动建议
action_items = []
if refund_rate > 0.15:
    action_items.append('①【紧急】逐单分析退款原因，建立退款分类标签（质量/物流/描述/其他）')
if has_backend and full_cvr < 0.02:
    action_items.append('②【重点】优化商详页主图视频+核心卖点首屏展示，提升访客→下单率')
if has_backend and pay_rate < 0.7:
    action_items.append('③【重要】配置下单未支付自动催付机制（优惠券/短信提醒）')
action_items.append(f'④【常规】关注竞品618后价格策略，合理调整促销力度')
action_items.append(f'⑤【常规】高客单客户建联与私域沉淀，推动复购与口碑转介绍')

insights.append({
    'level': 'blue', 'title': 'P0/P1 行动建议',
    'text': '；'.join(action_items)
})

# 填充至10条
placeholder_insights = [
    {'level': 'blue', 'title': '内容运营建议', 'text': '冰箱作为高决策品类，建议加强"真实安装案例+评测"内容矩阵，在商品详情页嵌入短视频与用户晒单。'},
    {'level': 'blue', 'title': '周期性关注', 'text': f'持续监控每日访客-下单-支付-成团漏斗各环节转化率变化，异常波动及时预警。统计周期：日报+周报+月报。'},
    {'level': 'green', 'title': '竞品态势', 'text': '建议每周扫描竞品店铺（海尔/美的/容声）在主销价位段的促销活动与新品上架情况，及时调整差异化策略。'},
    {'level': 'green', 'title': '服务体验', 'text': '高客单家电的物流时效与安装服务是影响好评率的关键因子，建议监控发货时效与安装预约完成率。'},
]
for ins in placeholder_insights:
    if len(insights) < 10:
        insights.append(ins)

print(f'生成 {len(insights)} 条洞察')

# ============================================================
# 第七部分：保存 Excel
# ============================================================
print(f'\n{"=" * 60}')
print('第七部分：保存数据')
print('=' * 60)

with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    # Overview
    pd.DataFrame({
        'date': [str(target_date)],
        'total_sales': [total_sales],
        'refund_amount': [refund_amount],
        'orders': [total_orders],
        'pay_users': [unique_buyers],
        'visitors': [funnel_visitors],
        'aov': [round(aov, 2)],
        'refund_rate': [round(refund_rate, 4)],
        'pay_cvr': [round(full_cvr, 4)],
    }).to_excel(writer, sheet_name='overview', index=False)

    # Funnel
    pd.DataFrame({
        'visitors': [funnel_visitors],
        'order_placers': [funnel_orders],
        'payers': [funnel_payers],
        'success_buyers': [funnel_success],
        'order_rate': [round(order_rate, 4)],
        'pay_rate': [round(pay_rate, 4)],
        'success_rate': [round(success_rate, 4)],
        'full_cvr': [round(full_cvr, 4)],
        'inquiry_count': [funnel_inquiry],
        'inquiry_rate': [round(inquiry_rate, 4)],
    }).to_excel(writer, sheet_name='funnel', index=False)

    # Category
    cat_out = category_agg[['category', 'sales_amount', 'qty', 'orders', 'refund_amount', 'refund_rate', 'sales_share', 'pay_users']].copy()
    cat_out.to_excel(writer, sheet_name='category', index=False)

    # Products
    prod_out = product_agg[['product_name', 'sku', 'sales_amount', 'qty', 'orders', 'pay_users', 'avg_price', 'refund_amount', 'refund_rate']].copy()
    prod_out.to_excel(writer, sheet_name='products', index=False)

    # Traffic
    if len(traffic_stats) > 0:
        traffic_stats.to_excel(writer, sheet_name='traffic', index=False)

    # Trend
    pd.DataFrame({
        'days': trend_days,
        'sales': trend_sales,
        'visitors': trend_visitors,
        'orders': trend_orders_list,
        'cvr': trend_cvr_list,
    }).to_excel(writer, sheet_name='trend', index=False)

    # Insights
    pd.DataFrame(insights).to_excel(writer, sheet_name='insights', index=False)

print(f'[OK] 数据已保存: {output_path}')

# ============================================================
# 打印概览
# ============================================================
print(f'\n{"=" * 60}')
print(f'========== 2026年6月18日 BI 日报概览 ==========')
print(f'  日期:             {target_date}')
print(f'  成团金额:         ¥{total_sales:,.2f}（订单数据为准）')
print(f'  退款金额:         ¥{refund_amount:,.2f}')
print(f'  成团订单数:       {total_orders}')
print(f'  客单价:           ¥{aov:,.2f}')
print(f'  退款率:           {refund_rate:.1%}')
if has_backend:
    print(f'  ---')
    print(f'  访客数(后台):     {funnel_visitors}')
    print(f'  全店转化率:       {full_cvr:.2%}')
    print(f'  漏斗: 访客{funnel_visitors}→下单{funnel_orders}→支付{funnel_payers}→成团{funnel_success}')
print(f'  品类数:           {len(category_agg)}')
print(f'  商品数:           {len(product_agg)}')
print(f'  流量省份数:       {len(traffic_stats)}')
print(f'  洞察建议数:       {len(insights)}')
print(f'==================================================')
