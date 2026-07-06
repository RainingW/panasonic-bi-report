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

def safe_text(value):
    if pd.isna(value):
        return ''
    return str(value)

def build_report_from_excel(excel_path):
    workbook = pd.read_excel(excel_path, sheet_name=None)

    # --- overview ---
    overview = {}
    if 'overview' in workbook:
        rows = workbook['overview'].fillna(0).to_dict(orient='records')
        overview = rows[0] if rows else {}
    if not overview:
        overview = {
            'date': datetime.now().strftime('%Y-%m-%d'),
            'total_sales': 0, 'refund_amount': 0,
            'orders': 0, 'pay_users': 0, 'visitors': 0,
            'aov': 0, 'refund_rate': 0, 'pay_cvr': 0,
        }

    # --- funnel ---
    funnel = {}
    if 'funnel' in workbook:
        rows = workbook['funnel'].fillna(0).to_dict(orient='records')
        funnel = rows[0] if rows else {}
    if not funnel:
        funnel = {
            'visitors': 0, 'order_placers': 0, 'payers': 0, 'success_buyers': 0,
            'order_rate': 0, 'pay_rate': 0, 'success_rate': 0, 'full_cvr': 0,
            'inquiry_count': 0, 'inquiry_rate': 0,
        }

    # --- category ---
    category = []
    if 'category' in workbook:
        df = workbook['category'].fillna(0)
        for _, row in df.iterrows():
            category.append({
                'category': safe_text(row.get('category', '')),
                'sales_amount': int(parse_number(row.get('sales_amount', 0))),
                'qty': int(parse_number(row.get('qty', 0))),
                'orders': int(parse_number(row.get('orders', 0))),
                'pay_users': int(parse_number(row.get('pay_users', 0))),
                'refund_amount': int(parse_number(row.get('refund_amount', 0))),
                'refund_rate': float(parse_number(row.get('refund_rate', 0))),
                'sales_share': float(parse_number(row.get('sales_share', 0))),
            })
        category.sort(key=lambda x: x['sales_amount'], reverse=True)

    # --- products ---
    products = []
    if 'products' in workbook:
        df = workbook['products'].fillna(0)
        for _, row in df.iterrows():
            products.append({
                'name': safe_text(row.get('product_name', '')),
                'sku': safe_text(row.get('sku', 'N/A')),
                'sales_amount': int(parse_number(row.get('sales_amount', 0))),
                'qty': int(parse_number(row.get('qty', 0))),
                'orders': int(parse_number(row.get('orders', 0))),
                'pay_users': int(parse_number(row.get('pay_users', 0))),
                'avg_price': int(parse_number(row.get('avg_price', 0))),
                'refund_amount': int(parse_number(row.get('refund_amount', 0))),
                'refund_rate': float(parse_number(row.get('refund_rate', 0))),
            })
        products.sort(key=lambda x: x['sales_amount'], reverse=True)
    if not products:
        products = [{'name': '暂无数据', 'sku': '-', 'sales_amount': 0, 'qty': 0,
                     'orders': 0, 'pay_users': 0, 'avg_price': 0,
                     'refund_amount': 0, 'refund_rate': 0}]

    # --- traffic ---
    traffic = []
    if 'traffic' in workbook:
        df = workbook['traffic'].fillna(0)
        for _, row in df.iterrows():
            traffic.append({
                'source': safe_text(row.get('source', '')),
                'value': int(parse_number(row.get('value', 0))),
                'sales': int(parse_number(row.get('sales', 0))),
            })

    # --- trend ---
    trend = {'days': [], 'sales': [], 'visitors': [], 'orders': [], 'cvr': []}
    if 'trend' in workbook:
        df = workbook['trend']
        trend['days'] = [safe_text(d) for d in df['days'].tolist()]
        trend['sales'] = [int(parse_number(s)) for s in df['sales'].tolist()]
        trend['visitors'] = [int(parse_number(v)) for v in df['visitors'].tolist()]
        trend['orders'] = [int(parse_number(o)) for o in df['orders'].tolist()]
        trend['cvr'] = [float(parse_number(c)) for c in df['cvr'].tolist()]

    # --- insights ---
    insights = []
    if 'insights' in workbook:
        df = workbook['insights']
        for _, row in df.iterrows():
            insights.append({
                'level': safe_text(row.get('level', 'blue')),
                'title': safe_text(row.get('title', '')),
                'text': safe_text(row.get('text', '')),
            })

    return {
        'overview': overview,
        'funnel': funnel,
        'category': category,
        'products': products,
        'traffic': traffic,
        'trend': trend,
        'insights': insights,
    }

def render_index_html(report):
    with open(TEMPLATE_PATH, 'r', encoding='utf-8') as f:
        template = f.read()
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
