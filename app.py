import os
import re
import json
import pandas as pd
from flask import Flask, render_template, request, send_file, jsonify
import Levenshtein
from io import BytesIO

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'outputs'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

VERSION = "1.0.0"

# 加载匹配数据
PROCESS_FIELDS = []  # 加工字段表

def load_match_data():
    global PROCESS_FIELDS
    
    local_file = os.path.join(os.path.dirname(__file__), 'templates', '加工字段表.xlsx')
    if os.path.exists(local_file):
        try:
            df = pd.read_excel(local_file, sheet_name='Sheet1')
            if '参数说明' in df.columns:
                PROCESS_FIELDS = df.to_dict('records')
                print(f"加工字段: {len(PROCESS_FIELDS)} 条")
        except Exception as e:
            print(f"加载失败: {e}")

load_match_data()

# 语义相关词映射
SEMANTIC_MAP = {
    '近': ['最近', '近几', '最近几'],
    '月': ['月度', '月份'],
    '年': ['年度', '年份'],
    '销售': ['营收', '收入', '生意'],
    '采购': ['进货', '供应', '购买'],
    '发票': ['票', '开票'],
    '税': ['税务', '纳税'],
    '客户': ['采购方', '买方', '购买方'],
    '供应商': ['供货方', '卖方'],
    '金额': ['额度', '数额'],
    '分析': ['分析', '评估'],
    '趋势': ['走势', '变化'],
}

def clean_text(s):
    if not s:
        return ""
    s = str(s).replace(' ', '').strip()
    return s

def get_semantic_score(user_field, target_field):
    score = 0
    for user_word, related_words in SEMANTIC_MAP.items():
        if user_word in str(user_field):
            for related in related_words:
                if related in str(target_field):
                    score += 20
    return score

def parse_excel_fields(filepath):
    """解析Excel文件中的字段"""
    fields = []
    try:
        df = pd.read_excel(filepath)
        if '字段名称' in df.columns:
            fields = df['字段名称'].dropna().astype(str).tolist()
            fields = [x.strip() for x in fields if x.strip()]
    except Exception as e:
        print(f"解析失败: {e}")
    return fields

def find_match(user_field):
    user_field = str(user_field).strip()
    if not user_field:
        return None
    
    user_clean = clean_text(user_field)
    best_match = None
    
    for row in PROCESS_FIELDS:
        target_cn = str(row.get('参数说明', '')).strip()  # 中文名
        target_en = str(row.get('参数名称', '')).strip()  # 英文名
        target_interface = str(row.get('接口', '')).strip()  # 接口名
        
        if not target_cn:
            continue
        
        target_clean = clean_text(target_cn)
        
        # 检查中文名匹配
        base_score = 0
        match_type = ''
        
        if user_field == target_cn or user_clean == target_clean:
            base_score = 100
            match_type = '完全匹配'
        else:
            try:
                sim = Levenshtein.ratio(user_clean, target_clean)
                if sim >= 0.4:
                    base_score = int(sim * 100)
                    match_type = '推荐'
            except:
                pass
        
        if base_score > 0:
            semantic_bonus = get_semantic_score(user_field, target_cn)
            total_score = min(100, base_score + semantic_bonus)
            
            if best_match is None or total_score > best_match['score']:
                best_match = {
                    'user_field': user_field,
                    'matched_cn': target_cn,
                    'matched_en': target_en,
                    'matched_interface': target_interface,
                    'source': target_interface,
                    'match_type': match_type,
                    'score': total_score
                }
    
    return best_match

@app.route('/')
def index():
    return render_template('index.html', version=VERSION)

@app.route('/template/excel')
def download_template():
    template_file = os.path.join(os.path.dirname(__file__), 'templates', '加工字段表.xlsx')
    return send_file(template_file, as_attachment=True, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.route('/match', methods=['POST'])
def match_fields():
    try:
        single_field = request.form.get('single_field')
        
        if single_field:
            user_fields = [single_field.strip()]
        elif 'file' not in request.files:
            return jsonify({'error': '请上传文件或输入字段'}), 400
        else:
            file = request.files['file']
            if file.filename == '':
                return jsonify({'error': '请选择文件'}), 400
            
            ext = os.path.splitext(file.filename)[1].lower()
            if ext not in ['.xlsx', '.xls']:
                return jsonify({'error': '只支持Excel文件'}), 400
            
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(filepath)
            
            user_fields = parse_excel_fields(filepath)
            if not user_fields:
                return jsonify({'error': '未能解析出字段'}), 400
        
        results = []
        for field in user_fields:
            result = find_match(field)
            if result:
                results.append(result)
            else:
                results.append({
                    'user_field': field,
                    'matched_cn': '-',
                    'matched_en': '-',
                    'matched_interface': '-',
                    'source': '-',
                    'match_type': '匹配不到',
                    'score': 0
                })
        
        total = len(results)
        exact = len([r for r in results if r['match_type'] == '完全匹配'])
        recommend = len([r for r in results if r['match_type'] == '推荐'])
        failed = len([r for r in results if r['match_type'] == '匹配不到'])
        
        result_df = pd.DataFrame(results)
        result_df.to_excel(os.path.join(app.config['OUTPUT_FOLDER'], 'result.xlsx'), index=False)
        
        return jsonify({
            'success': True,
            'stats': {'total': total, 'exact': exact, 'recommend': recommend, 'failed': failed},
            'results': results[:100]
        })
    
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/download')
def download_result():
    result_path = os.path.join(app.config['OUTPUT_FOLDER'], 'result.xlsx')
    if os.path.exists(result_path):
        return send_file(result_path, as_attachment=True)
    return "文件未找到", 404

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)