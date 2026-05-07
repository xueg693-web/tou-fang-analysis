import os
import json
import math
import urllib.request
import pandas as pd
from flask import Flask, request, jsonify, render_template, send_file
from werkzeug.utils import secure_filename
import openpyxl
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import PatternFill, Font
from datetime import datetime

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
EXPORT_FOLDER = os.path.join(BASE_DIR, 'exports')
CONFIG_FILE = os.path.join(BASE_DIR, 'config', 'metrics.json')

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(EXPORT_FOLDER, exist_ok=True)
os.makedirs(os.path.join(BASE_DIR, 'config'), exist_ok=True)

# ──────────────────────────────────────────────────────────────────────────────
# 工具函数
# ──────────────────────────────────────────────────────────────────────────────

TIME_COLUMN_KEYWORDS = ['week', 'date', 'month', 'period', 'time', 'day',
                        '周', '日期', '月份', '时间', '日']

def detect_time_column(columns):
    """自动检测时间/周期列"""
    for col in columns:
        if any(kw in col.lower() for kw in TIME_COLUMN_KEYWORDS):
            return col
    return columns[0] if columns else None

def load_metrics_config():
    """加载指标配置"""
    if not os.path.exists(CONFIG_FILE):
        return []
    try:
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except (json.JSONDecodeError, IOError):
        return []

def save_metrics_config(configs):
    """保存指标配置"""
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(configs, f, ensure_ascii=False, indent=2)
    except IOError as e:
        raise IOError(f'保存配置失败: {str(e)}')

def calculate_metrics(df, metric_defs):
    """
    计算自定义指标，返回新列字典。
    metric_defs: [{name, numerator, denominator, type, isFormula?, formula?, currency?, unit?}, ...]
    """
    result_cols = {}
    for m in metric_defs:
        name = m['name']
        dtype = m['type']

        if m.get('isFormula') and m.get('formula'):
            raw = _eval_formula_series(df, m['formula'])
            if raw is None:
                continue
        else:
            num = m.get('numerator', '')
            den = m.get('denominator', '')
            if num not in df.columns or den not in df.columns:
                continue
            raw = df[num] / df[den].replace(0, float('nan'))
            if dtype == 'percent':
                raw = raw * 100

        result_cols[name] = pd.to_numeric(raw, errors='coerce').round(4)
    return result_cols


def _eval_formula_series(df, formula):
    """
    安全地对 DataFrame 整列求值公式。
    公式中字段名用 {字段名} 包裹，支持 + - * / ( ) 和数字。
    """
    import re
    fields = re.findall(r'\{([^}]+)\}', formula)
    for f in fields:
        if f not in df.columns:
            return None
    # 将 {字段名} 替换为 col_N 变量名，避免字段名含特殊字符
    local_vars = {}
    expr = formula
    for i, f in enumerate(fields):
        var = f'col_{i}'
        expr = expr.replace('{' + f + '}', var)
        local_vars[var] = pd.to_numeric(df[f], errors='coerce')
    # 只允许安全字符：数字、运算符、括号、空格、col_N 变量名
    if not re.match(r'^[\d\s\+\-\*/\(\)\.col_0-9]+$', expr):
        return None
    try:
        return eval(expr, {"__builtins__": {}}, local_vars)  # noqa: S307
    except Exception:
        return None

def detect_anomalies(series, threshold=0.5):
    """环比变化超过阈值视为异常，返回布尔 Series"""
    numeric = pd.to_numeric(series, errors='coerce')
    pct_change = numeric.pct_change().abs()
    return pct_change.fillna(0) > threshold

def format_value(val, dtype, currency=''):
    """格式化单个数值用于前端显示"""
    if val is None or (isinstance(val, float) and math.isnan(val)):
        return '-'
    if dtype == 'percent':
        return f"{val:.2f}%"
    if dtype == 'currency':
        return f"{currency}{val:.2f}"
    if dtype == 'count':
        return f"{int(val):,}"
    return str(val)

# ──────────────────────────────────────────────────────────────────────────────
# 路由
# ──────────────────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/upload', methods=['POST'])
def upload_csv():
    """上传一个或多个 CSV，返回字段列表和预览数据"""
    files = request.files.getlist('files')
    if not files:
        return jsonify({'error': '未选择文件'}), 400

    all_data = []
    for f in files:
        filename = secure_filename(f.filename)
        if not filename.endswith('.csv'):
            continue
        save_path = os.path.join(UPLOAD_FOLDER, filename)
        f.save(save_path)

        try:
            df = pd.read_csv(save_path)
        except Exception as e:
            return jsonify({'error': f'解析 {filename} 失败: {str(e)}'}), 400

        time_col = detect_time_column(list(df.columns))
        all_data.append({
            'filename': filename,
            'platform': filename.replace('.csv', ''),
            'columns': list(df.columns),
            'time_column': time_col,
            'rows': len(df),
            'preview': df.head(3).fillna('').to_dict(orient='records'),
        })

    if not all_data:
        return jsonify({'error': '未找到有效的 CSV 文件'}), 400

    return jsonify({'files': all_data})

@app.route('/api/analyze', methods=['POST'])
def analyze():
    """执行分析：读取已上传 CSV + 应用指标配置，返回图表数据和表格数据"""
    payload = request.get_json()
    files_info = payload.get('files', [])        # [{filename, platform, time_column}, ...]
    metric_defs = payload.get('metrics', [])     # 来自前端的指标定义列表

    if not files_info:
        return jsonify({'error': '请先上传 CSV 文件'}), 400

    charts_data = {}   # {metric_name: {labels:[], datasets:[{label, data:[]}]}}
    tables_data = []   # [{platform, rows:[{col: val}], anomalies:{col: [idx]}}]

    for fi in files_info:
        filename = fi['filename']
        platform = fi.get('platform', filename)
        time_col = fi.get('time_column')

        csv_path = os.path.join(UPLOAD_FOLDER, secure_filename(filename))
        if not os.path.exists(csv_path):
            continue

        try:
            df = pd.read_csv(csv_path)
        except Exception as e:
            continue
        if time_col and time_col in df.columns:
            labels = df[time_col].astype(str).tolist()
        else:
            labels = [str(i) for i in range(len(df))]

        # 计算指标新列
        extra_cols = calculate_metrics(df, metric_defs)
        for col_name, series in extra_cols.items():
            df[col_name] = series

        # 构建图表数据（原始数值列 + 计算列）
        for m in metric_defs:
            mname = m['name']
            if mname not in df.columns:
                continue
            if mname not in charts_data:
                charts_data[mname] = {
                    'labels': labels,
                    'datasets': [],
                    'type': m['type'],
                    'currency': m.get('currency', ''),
                    'unit': m.get('unit', ''),
                }
            charts_data[mname]['datasets'].append({
                'label': f"{platform}",
                'data': [None if (isinstance(v, float) and math.isnan(v)) else round(v, 4)
                         for v in df[mname].tolist()],
            })

        # 检测异常
        anomalies = {}
        all_metric_cols = [m['name'] for m in metric_defs if m['name'] in df.columns]
        for col in all_metric_cols:
            anom_mask = detect_anomalies(df[col])
            anomalies[col] = [i for i, v in enumerate(anom_mask) if v]

        # 构建表格数据
        display_cols = ([time_col] if time_col else []) + \
                       [c for c in df.columns if c != time_col]
        table_rows = df[display_cols].fillna('').to_dict(orient='records')
        tables_data.append({
            'platform': platform,
            'columns': display_cols,
            'rows': table_rows,
            'anomalies': anomalies,
        })

    return jsonify({
        'charts': charts_data,
        'tables': tables_data,
    })

@app.route('/api/metrics', methods=['GET'])
def get_metrics_configs():
    """获取所有已保存的指标配置组"""
    return jsonify(load_metrics_config())

@app.route('/api/metrics', methods=['POST'])
def save_metrics():
    """保存一个新的指标配置组"""
    data = request.get_json()
    name = data.get('name', '').strip()
    metrics = data.get('metrics', [])
    if not name:
        return jsonify({'error': '配置名称不能为空'}), 400
    configs = load_metrics_config()
    # 同名则覆盖
    configs = [c for c in configs if c['name'] != name]
    configs.append({'name': name, 'metrics': metrics})
    save_metrics_config(configs)
    return jsonify({'success': True})

@app.route('/api/metrics/<name>', methods=['DELETE'])
def delete_metrics(name):
    """删除指定名称的指标配置"""
    configs = load_metrics_config()
    configs = [c for c in configs if c['name'] != name]
    save_metrics_config(configs)
    return jsonify({'success': True})

@app.route('/api/export/excel', methods=['POST'])
def export_excel():
    """导出 Excel：每个平台一个 Sheet，包含格式化数据"""
    payload = request.get_json()
    tables_data = payload.get('tables', [])
    metric_defs = payload.get('metrics', [])
    config_name = payload.get('config_name', 'export')

    dtype_map = {m['name']: m for m in metric_defs}

    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # 删除默认空 sheet

    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    anomaly_fill = PatternFill(start_color="FF6B6B", end_color="FF6B6B", fill_type="solid")

    used_sheet_names = set()
    for table in tables_data:
        platform_raw = table['platform'][:31]
        platform = platform_raw
        suffix = 1
        while platform in used_sheet_names:
            platform = f"{platform_raw[:28]}_{suffix}"
            suffix += 1
        used_sheet_names.add(platform)
        ws = wb.create_sheet(title=platform)

        columns = table['columns']
        rows = table['rows']
        anomalies = table.get('anomalies', {})

        # 写表头
        for ci, col in enumerate(columns, 1):
            cell = ws.cell(row=1, column=ci, value=col)
            cell.fill = header_fill
            cell.font = header_font

        # 写数据行
        for ri, row in enumerate(rows, 2):
            for ci, col in enumerate(columns, 1):
                val = row.get(col, '')
                # 格式化显示
                if col in dtype_map:
                    m = dtype_map[col]
                    try:
                        numeric_val = float(val)
                        cell = ws.cell(row=ri, column=ci, value=numeric_val)
                        if m['type'] == 'percent':
                            cell.number_format = '0.00"%"'
                        elif m['type'] == 'currency':
                            sym = m.get('currency', '')
                            cell.number_format = f'"{sym}"#,##0.00'
                        elif m['type'] == 'count':
                            cell.number_format = '#,##0'
                    except (ValueError, TypeError):
                        ws.cell(row=ri, column=ci, value=val)
                else:
                    ws.cell(row=ri, column=ci, value=val)

                # 标红异常
                if col in anomalies and (ri - 2) in anomalies[col]:
                    ws.cell(row=ri, column=ci).fill = anomaly_fill

        # 自适应列宽
        for col_cells in ws.columns:
            max_len = max((len(str(c.value)) for c in col_cells if c.value), default=10)
            ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 4, 30)

    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    safe_name = config_name.replace('/', '_').replace('\\', '_')
    filename = f"{timestamp}_{safe_name}.xlsx"
    export_path = os.path.join(EXPORT_FOLDER, filename)
    try:
        wb.save(export_path)
    except Exception as e:
        return jsonify({'error': f'Excel 导出失败: {str(e)}'}), 500

    return send_file(export_path, as_attachment=True, download_name=filename)

@app.route('/api/ai-analyze', methods=['POST'])
def ai_analyze():
    """后端代理调用 Anthropic API，避免浏览器 CORS 限制"""
    payload = request.get_json()
    api_key = payload.get('api_key', '').strip()
    data_summary = payload.get('data_summary', '')

    if not api_key:
        return jsonify({'error': '请填写 Anthropic API Key'}), 400
    if not data_summary:
        return jsonify({'error': '数据摘要为空'}), 400

    system_prompt = (
        "你是一位拥有10年经验的资深投放数据优化师，专注于效果广告投放优化，"
        "熟悉各大流量平台（字节、腾讯、百度、Meta、Google等）的算法机制和优化策略。\n"
        "根据用户提供的投放数据，给出专业、深度、可落地的分析报告。\n\n"
        "报告结构（严格按此格式输出）：\n"
        "## 一、数据核心发现\n列出3-5个关键数据洞察，每条需有具体数值支撑\n\n"
        "## 二、问题诊断\n识别表现异常的指标，分析可能原因\n\n"
        "## 三、优化建议\n给出5-8条具体可执行建议，每条标注优先级：🔴紧急 / 🟡重要 / 🟢建议\n\n"
        "## 四、平台对比结论\n（如有多平台数据）指出哪个平台表现更优及原因；单平台则分析趋势\n\n"
        "## 五、下阶段策略\n给出简明的下阶段投放策略方向\n\n"
        "语言要求：专业但易懂，用数据说话，避免空话套话。"
    )

    body = json.dumps({
        'model': 'claude-sonnet-4-5',
        'max_tokens': 2048,
        'system': system_prompt,
        'messages': [{'role': 'user', 'content': f'请分析以下投放数据并给出专业建议：\n\n{data_summary}'}]
    }).encode('utf-8')

    req = urllib.request.Request(
        'https://api.anthropic.com/v1/messages',
        data=body,
        headers={
            'Content-Type': 'application/json',
            'x-api-key': api_key,
            'anthropic-version': '2023-06-01',
        },
        method='POST'
    )

    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            result = json.loads(resp.read().decode('utf-8'))
        text = ''.join(b.get('text', '') for b in result.get('content', []))
        return jsonify({'result': text})
    except urllib.error.HTTPError as e:
        err_body = e.read().decode('utf-8', errors='replace')
        try:
            err_json = json.loads(err_body)
            msg = err_json.get('error', {}).get('message', err_body)
        except Exception:
            msg = err_body
        return jsonify({'error': f'API 错误 {e.code}: {msg}'}), 502
    except Exception as e:
        return jsonify({'error': f'请求失败: {str(e)}'}), 502

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5001))
    print("=" * 50)
    print("投放数据分析工具已启动")
    print(f"请在浏览器访问: http://127.0.0.1:{port}")
    print("=" * 50)
    app.run(debug=False, port=port, host='0.0.0.0')
