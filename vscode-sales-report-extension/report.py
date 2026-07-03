import os
import sys
import json
import pandas as pd
from datetime import datetime

WORKDIR = os.path.dirname(__file__)

TEMPLATE_PATH = os.path.join(WORKDIR, 'index_template.html')
OUTPUT_PATH = os.path.join(WORKDIR, 'index.html')

def parse_number(value):
    if pd.isna(value):
        return 0
    if isinstance(value, str):
        text = value.strip().replace('%', '')
        try:
            return float(text)
        except ValueError:
            return 0
    try:
        return float(value)
    except Exception:
        return 0

def parse_percent(value):
    if pd.isna(value):
        return 0
    if isinstance(value, str) and '%' in value:
        return parse_number(value) / 100
    return parse_number(value)

def safe_text(value):
    if pd.isna(value):
        return ''
    return str(value)

def normalize_raw_order_df(df):
    if df is None or df.empty:
        return df
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    if '订单状态' in df.columns:
        df['order_status'] = df['订单状态'].astype(str)
    else:
        df['order_status'] = df.get('order_status', '').astype(str)
    if '支付时间' in df.columns:
        df['payment_date'] = pd.to_datetime(df['支付时间'], errors='coerce')
    if '用户支付金额' in df.columns:
        df['user_payment_amount'] = pd.to_numeric(df['用户支付金额'], errors='coerce').fillna(0)
    elif '用户实付金额(元)' in df.columns:
        df['user_payment_amount'] = pd.to_numeric(df['用户实付金额(元)'], errors='coerce').fillna(0)
    else:
        df['user_payment_amount'] = 0
    buyer_col = None
    for col in ['用户购买手机号', '消费者资料', 'buyer_id']:
        if col in df.columns:
            buyer_col = col
            break
    if buyer_col:
        df['buyer_id'] = df[buyer_col].astype(str).str.strip()
        df.loc[df['buyer_id'] == '', 'buyer_id'] = pd.NA
    else:
        df['buyer_id'] = df['订单号'].astype(str).where(df['订单号'].notna(), pd.NA)
    df['is_canceled'] = df['order_status'].str.contains('已取消', na=False)
    df['is_refunded'] = df['order_status'].str.contains('退款成功', na=False)
    df['is_active'] = ~df['is_canceled'] & ~df['is_refunded']
    return df

def load_template():
    with open(TEMPLATE_PATH, 'r', encoding='utf-8') as f:
        return f.read()

def build_overview(raw_df):
    if raw_df is None or raw_df.empty:
        return {
            'date': datetime.now().strftime('%Y-%m-%d'),
            'total_sales': 0,
            'payment_amount': 0,
            'refund_amount': 0,
            'orders': 0,
            'pay_cvr': 0,
            'visitors': 0,
            'aov': 0,
            'refund_rate': 0,
            'add_to_cart': 0,
            'favorites': 0
        }
    df = normalize_raw_order_df(raw_df) if '订单状态' in raw_df.columns or '支付时间' in raw_df.columns else raw_df.copy()
    active = df[df['is_active']] if 'is_active' in df.columns else df
    refund = df[df['is_refunded']] if 'is_refunded' in df.columns else df.iloc[0:0]
    total_sales = int(active['user_payment_amount'].sum()) if 'user_payment_amount' in active else 0
    payment_amount = total_sales
    refund_amount = int(refund['user_payment_amount'].sum()) if 'user_payment_amount' in refund else 0
    orders = active['订单号'].nunique() if '订单号' in active else len(active)
    if 'buyer_id' in active.columns:
        pay_users = active['buyer_id'].dropna().nunique()
        if pay_users == 0:
            pay_users = orders  # 买家标识全为空时回退到订单数
    else:
        pay_users = orders
    visitors = int(active['visitors'].sum()) if 'visitors' in active else int(pay_users)
    pay_cvr = payment_amount / visitors if visitors else 0
    aov = total_sales / pay_users if pay_users else 0
    refund_rate = refund_amount / (total_sales + refund_amount) if (total_sales + refund_amount) else 0
    date = df['payment_date'].max() if 'payment_date' in df else (df['date'].max() if 'date' in df.columns else datetime.now())
    return {
        'date': safe_text(date),
        'total_sales': int(total_sales),
        'payment_amount': int(payment_amount),
        'refund_amount': int(refund_amount),
        'orders': int(orders),
        'pay_cvr': round(pay_cvr, 4),
        'visitors': int(visitors),
        'aov': round(aov, 2),
        'refund_rate': round(refund_rate, 4),
        'add_to_cart': int(active['cart_add'].sum()) if 'cart_add' in active else 0,
        'favorites': int(active['favorites'].sum()) if 'favorites' in active else 0
    }

def build_category(df):
    if df is None or df.empty:
        return []
    rows = []
    for _, row in df.iterrows():
        rows.append({
            'category': safe_text(row.get('category') or row.get('品类') or row.get('product_category')),
            'sales_amount': int(parse_number(row.get('sales_amount') or row.get('销售额'))),
            'qty': int(parse_number(row.get('qty') or row.get('销量'))),
            'pay_users': int(parse_number(row.get('pay_users') or row.get('支付人数'))),
            'visitors': int(parse_number(row.get('visitors') or row.get('访客'))),
            'cvr': parse_percent(row.get('cvr') or row.get('支付转化率')),
            'refund_rate': parse_percent(row.get('refund_rate') or row.get('退款率')),
            'favorites': int(parse_number(row.get('favorites') or row.get('收藏'))),
            'cart_add': int(parse_number(row.get('cart_add') or row.get('加购'))),
            'new_user_rate': parse_percent(row.get('new_user_rate') or row.get('新客率') or row.get('新客%')),
            'repeat_user_rate': parse_percent(row.get('repeat_user_rate') or row.get('老客率') or row.get('老客%'))
        })
    return rows

def build_products(df):
    if df is None or df.empty:
        return []
    rows = []
    for _, row in df.iterrows():
        rows.append({
            'name': safe_text(row.get('name') or row.get('product_name') or row.get('商品名称')),
            'sku': safe_text(row.get('sku') or row.get('SKU')),
            'sales_amount': int(parse_number(row.get('sales_amount') or row.get('销售额'))),
            'qty': int(parse_number(row.get('qty') or row.get('销量'))),
            'pay_users': int(parse_number(row.get('pay_users') or row.get('支付人数'))),
            'visitors': int(parse_number(row.get('visitors') or row.get('访客'))),
            'cvr': parse_percent(row.get('cvr') or row.get('支付转化率')),
            'refund_rate': parse_percent(row.get('refund_rate') or row.get('退款率')),
            'cart_add': int(parse_number(row.get('cart_add') or row.get('加购'))),
            'favorites': int(parse_number(row.get('favorites') or row.get('收藏'))),
            'avg_price': int(parse_number(row.get('avg_price') or row.get('客单价'))),
            'stay_time': int(parse_number(row.get('stay_time') or row.get('停留时间'))),
            'bounce_rate': parse_percent(row.get('bounce_rate') or row.get('跳失率')),
            'new_user_rate': parse_percent(row.get('new_user_rate') or row.get('新客率') or row.get('新客%')),
            'repeat_user_rate': parse_percent(row.get('repeat_user_rate') or row.get('老客率') or row.get('老客%'))
        })
    return rows

def build_trend(df):
    if df is None or df.empty:
        return {'days': [], 'sales': [], 'visitors': []}
    df = normalize_raw_order_df(df) if '支付时间' in df.columns or '订单状态' in df.columns else df.copy()
    if 'payment_date' in df.columns:
        df['date'] = pd.to_datetime(df['payment_date'], errors='coerce')
    elif 'date' in df.columns:
        df['date'] = pd.to_datetime(df['date'], errors='coerce')
    else:
        df['date'] = pd.NaT
    df = df.dropna(subset=['date'])
    df['date'] = df['date'].dt.strftime('%Y-%m-%d')
    if 'is_active' in df.columns:
        df = df[df['is_active']]
    sales_field = 'user_payment_amount' if 'user_payment_amount' in df.columns else ('sales_amount' if 'sales_amount' in df.columns else df.columns[1] if len(df.columns) > 1 else df.columns[0])
    if 'visitors' not in df.columns:
        df['visitors'] = 1
    grouped = df.groupby('date').agg({sales_field:'sum', 'visitors':'sum'}).reset_index()
    return {
        'days': grouped['date'].tolist(),
        'sales': grouped[sales_field].astype(int).tolist(),
        'visitors': grouped['visitors'].astype(int).tolist()
    }

def build_traffic(df):
    if df is None or df.empty:
        return []
    df = normalize_raw_order_df(df) if '省' in df.columns or '订单状态' in df.columns else df.copy()
    if 'is_active' in df.columns:
        df = df[df['is_active']]
    if '省' in df.columns:
        grouped = df.groupby('省').agg({'订单号':'nunique'}).reset_index()
        rows = []
        for _, row in grouped.iterrows():
            rows.append({'source': safe_text(row['省'] or '未知'), 'value': int(row['订单号'])})
        return rows
    rows = []
    for _, row in df.iterrows():
        rows.append({'source': safe_text(row.get('source') or row.get('来源')), 'value': int(parse_number(row.get('value') or row.get('count') or row.get('访问量') or 0))})
    return rows

def build_funnels(df):
    if df is None or df.empty:
        return {'exposure': 0, 'visitor': 0, 'detail_page': 0, 'add_cart': 0, 'order': 0, 'payment': 0}
    row = df.iloc[0]
    return {
        'exposure': int(parse_number(row.get('exposure') or row.get('曝光'))),
        'visitor': int(parse_number(row.get('visitor') or row.get('访客'))),
        'detail_page': int(parse_number(row.get('detail_page') or row.get('详情页'))),
        'add_cart': int(parse_number(row.get('add_cart') or row.get('加购'))),
        'order': int(parse_number(row.get('order') or row.get('下单'))),
        'payment': int(parse_number(row.get('payment') or row.get('支付')))
    }

def derive_category_from_raw(df):
    if df is None or df.empty:
        return []
    df = normalize_raw_order_df(df)
    # 强制以商品三级类目做分类
    key = '商品三级类目' if '商品三级类目' in df.columns else None
    if not key:
        for field in ['商品二级类目', '商品一级类目', 'category', 'product_category', '品类']:
            if field in df.columns:
                key = field
                break
    if not key:
        return []
    active = df[df['is_active']] if 'is_active' in df.columns else df
    active[key] = active[key].fillna('未知')
    grouped = active.groupby(key).agg({
        'user_payment_amount':'sum',
        '商品数量(件)':'sum',
        '订单号':'nunique',
        'buyer_id':'nunique'
    }).reset_index().fillna(0)
    rows = []
    for _, row in grouped.iterrows():
        rows.append({
            'category': safe_text(row[key]),
            'sales_amount': int(parse_number(row.get('user_payment_amount') or 0)),
            'qty': int(parse_number(row.get('商品数量(件)') or 0)),
            'pay_users': int(parse_number(row.get('buyer_id') or 0)),
            'visitors': int(parse_number(row.get('订单号') or 0)),
            'cvr': 0,
            'refund_rate': 0,
            'favorites': 0,
            'cart_add': 0,
            'new_user_rate': 0,
            'repeat_user_rate': 0
        })
    return rows

def derive_products_from_raw(df):
    if df is None or df.empty:
        return []
    df = normalize_raw_order_df(df)
    key = None
    for field in ['商品', '商品名称', '货品名称', 'product_name', 'name']:
        if field in df.columns:
            key = field
            break
    if not key:
        return []
    active = df[df['is_active']] if 'is_active' in df.columns else df
    active[key] = active[key].fillna('未知')
    grouped = active.groupby(key).agg({
        'user_payment_amount':'sum',
        '商品数量(件)':'sum',
        '订单号':'nunique',
        'buyer_id':'nunique'
    }).reset_index().fillna(0)
    rows = []
    for _, row in grouped.iterrows():
        qty = int(parse_number(row.get('商品数量(件)') or 0))
        sales = int(parse_number(row.get('user_payment_amount') or 0))
        rows.append({
            'name': safe_text(row[key]),
            'sku': safe_text(row.get('商品id') or row.get('样式ID') or ''),
            'sales_amount': sales,
            'qty': qty,
            'pay_users': int(parse_number(row.get('buyer_id') or 0)),
            'visitors': int(parse_number(row.get('订单号') or 0)),
            'cvr': 0,
            'refund_rate': 0,
            'cart_add': 0,
            'favorites': 0,
            'avg_price': int(sales / max(1, qty)),
            'stay_time': 0,
            'bounce_rate': 0,
            'new_user_rate': 0,
            'repeat_user_rate': 0
        })
    return rows

def build_report_from_excel(excel_path):
    workbook = pd.read_excel(excel_path, sheet_name=None)
    report = {
        'overview': {},
        'category': [],
        'products': [],
        'trend': {'days': [], 'sales': [], 'visitors': []},
        'traffic': [],
        'funnels': {'exposure': 0, 'visitor': 0, 'detail_page': 0, 'add_cart': 0, 'order': 0, 'payment': 0}
    }

    if 'overview' in workbook:
        rows = workbook['overview'].fillna(0).to_dict(orient='records')
        report['overview'] = rows[0] if rows else report['overview']
    if 'category' in workbook:
        report['category'] = build_category(workbook['category'])
    if 'products' in workbook:
        report['products'] = build_products(workbook['products'])
    if 'trend' in workbook:
        report['trend'] = build_trend(workbook['trend'])
    if 'traffic' in workbook:
        report['traffic'] = build_traffic(workbook['traffic'])
    if 'funnels' in workbook:
        report['funnels'] = build_funnels(workbook['funnels'])

    if not report['overview'] or report['overview'] == {}:
        raw_sheet = list(workbook.values())[0]
        report['overview'] = build_overview(raw_sheet)

    raw_sheet = list(workbook.values())[0]
    if not report['category']:
        report['category'] = derive_category_from_raw(raw_sheet)

    if not report['products']:
        report['products'] = derive_products_from_raw(raw_sheet)

    if not report['trend']['days']:
        if 'trend' in workbook:
            report['trend'] = build_trend(workbook['trend'])
        else:
            report['trend'] = build_trend(raw_sheet)

    if not report['traffic']:
        if 'traffic' in workbook:
            report['traffic'] = build_traffic(workbook['traffic'])
        else:
            report['traffic'] = build_traffic(raw_sheet)

    if not report['funnels']['exposure']:
        report['funnels'] = {
            'exposure': 15000,
            'visitor': report['overview'].get('visitors', 0),
            'detail_page': max(0, int(report['overview'].get('visitors', 0) * 0.8)),
            'add_cart': report['overview'].get('add_to_cart', 0),
            'order': report['overview'].get('orders', 0),
            'payment': int(report['overview'].get('payment_amount', report['overview'].get('total_sales', 0)) / max(1, report['overview'].get('aov', 1)))
        }

    if not report['products']:
        report['products'] = [{'name': '样例商品', 'sku': 'SKU-001', 'sales_amount': 0, 'qty': 0, 'pay_users': 0, 'visitors': 0, 'cvr': 0, 'refund_rate':0, 'cart_add':0, 'favorites':0, 'avg_price':0,'stay_time':0,'bounce_rate':0,'new_user_rate':0,'repeat_user_rate':0}]

    if not report['category']:
        report['category'] = [{'category':'样例品类','sales_amount':0,'qty':0,'pay_users':0,'visitors':0,'cvr':0,'refund_rate':0,'favorites':0,'cart_add':0,'new_user_rate':0,'repeat_user_rate':0}]

    report['products'] = sorted(report['products'], key=lambda x: x.get('sales_amount',0), reverse=True)
    return report

def render_index_html(report):
    template = load_template()
    json_payload = json.dumps(report, ensure_ascii=False, default=str)
    html = template.replace('__REPORT_DATA__', json_payload)
    with open(OUTPUT_PATH, 'w', encoding='utf-8') as f:
        f.write(html)
    return OUTPUT_PATH

if __name__ == '__main__':
    print('Sales Report Generator (Python) - 生成 BI 显示页面')
    excel_path = None
    if len(sys.argv) > 1:
        excel_path = sys.argv[1].strip()
    if not excel_path:
        excel_path = input('Enter path to Excel file (relative to workspace or absolute): ').strip()
    if not excel_path:
        print('No path provided, exiting.')
        sys.exit(0)
    if not os.path.isabs(excel_path):
        excel_path = os.path.join(WORKDIR, excel_path)
    if not os.path.exists(excel_path):
        print(f'File not found: {excel_path}')
        sys.exit(1)

    report = build_report_from_excel(excel_path)
    output = render_index_html(report)
    print(f'Generated HTML dashboard: {output}')
    try:
        if os.name == 'nt':
            os.startfile(output)
    except Exception:
        pass
