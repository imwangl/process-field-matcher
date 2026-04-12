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
CUSTOM_FILE_PATH = None  # 用户上传的自定义匹配文件路径

def load_match_data_from_excel(filepath, sheet_name=None):
    """从Excel文件加载匹配数据"""
    fields = []
    try:
        if sheet_name:
            df = pd.read_excel(filepath, sheet_name=sheet_name)
            fields = df.to_dict('records')
        else:
            # 读取所有sheet
            xl = pd.ExcelFile(filepath)
            for sheet in xl.sheet_names:
                df = pd.read_excel(filepath, sheet_name=sheet)
                fields.extend(df.to_dict('records'))
        return fields
    except Exception as e:
        print(f"加载失败: {e}")
        return []

def load_match_data(custom_file=None):
    global PROCESS_FIELDS, CUSTOM_FILE_PATH
    
    # 优先使用用户上传的文件
    if custom_file and os.path.exists(custom_file):
        try:
            # 读取sheet1和sheet2
            xl = pd.ExcelFile(custom_file)
            all_fields = []
            
            for sheet in xl.sheet_names:
                df = pd.read_excel(custom_file, sheet_name=sheet)
                all_fields.extend(df.to_dict('records'))
            
            if all_fields:
                PROCESS_FIELDS = all_fields
                CUSTOM_FILE_PATH = custom_file
                print(f"使用自定义文件: {custom_file}, 总字段: {len(PROCESS_FIELDS)} 条")
                return
        except Exception as e:
            print(f"加载自定义文件失败: {e}, 尝试加载默认文件")
    
    # 回退到默认文件
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
        # 尝试读取所有sheet
        xl = pd.ExcelFile(filepath)
        
        # 常见列名
        possible_cols = ['字段名称', '字段', '名称', 'name', 'field', 'Field', '字段名']
        
        for sheet in xl.sheet_names:
            df = pd.read_excel(filepath, sheet_name=sheet)
            
            # 尝试找到包含字段名的列
            for col in df.columns:
                col_str = str(col).strip()
                if col_str in possible_cols:
                    # 找到匹配的列
                    fields = df[col].dropna().astype(str).tolist()
                    fields = [x.strip() for x in fields if x.strip()]
                    print(f"从sheet '{sheet}' 列 '{col}' 解析出 {len(fields)} 个字段")
                    return fields
            
            # 如果没找到，尝试第一列
            if len(df.columns) > 0:
                first_col = df.columns[0]
                fields = df[first_col].dropna().astype(str).tolist()
                fields = [x.strip() for x in fields if x.strip()]
                if fields:
                    print(f"从sheet '{sheet}' 第一列解析出 {len(fields)} 个字段")
                    return fields
    
    except Exception as e:
        print(f"解析失败: {e}")
    
    return fields

def is_english_field(field):
    """判断是否为英文字段（主要包含英文字母）"""
    if not field:
        return False
    field = str(field).strip()
    # 统计英文字母数量
    letter_count = sum(1 for c in field if c.isalpha() and c.isascii())
    # 如果超过50%是英文字母，认为是英文字段
    return letter_count > len(field) * 0.5

def find_match(user_field):
    user_field = str(user_field).strip()
    if not user_field:
        return None
    
    user_clean = clean_text(user_field)
    best_match = None
    exact_match = None  # 完全匹配
    
    is_english = is_english_field(user_field)
    
    for row in PROCESS_FIELDS:
        # 列2: 参数名称 (英文名)
        # 列3: 参数说明 (中文名)
        target_en = str(row.get('参数名称', row.get('参数名称', ''))).strip()  # 列2
        target_cn = str(row.get('参数说明', row.get('参数说明', ''))).strip()  # 列3
        target_interface = str(row.get('接口', row.get('接口', ''))).strip()
        
        base_score = 0
        match_type = ''
        matched_value = None
        
        if is_english:
            # 英文字段：优先匹配列2（参数名称），不匹配列3
            if target_en and (user_field == target_en or user_clean == clean_text(target_en)):
                base_score = 100
                match_type = '完全匹配'
                matched_value = target_en
            elif target_en:
                try:
                    sim = Levenshtein.ratio(user_clean, clean_text(target_en))
                    if sim >= 0.4:
                        base_score = int(sim * 100)
                        match_type = '推荐'
                        matched_value = target_en
                except:
                    pass
        else:
            # 中文字段：匹配列3（参数说明）
            if target_cn and (user_field == target_cn or user_clean == clean_text(target_cn)):
                base_score = 100
                match_type = '完全匹配'
                matched_value = target_cn
            elif target_cn:
                try:
                    sim = Levenshtein.ratio(user_clean, clean_text(target_cn))
                    if sim >= 0.4:
                        base_score = int(sim * 100)
                        match_type = '推荐'
                        matched_value = target_cn
                except:
                    pass
        
        if base_score > 0:
            # 对于英文，增加语义得分
            if is_english and target_en:
                semantic_bonus = get_semantic_score(user_field, target_en)
            elif not is_english and target_cn:
                semantic_bonus = get_semantic_score(user_field, target_cn)
            else:
                semantic_bonus = 0
                
            total_score = min(100, base_score + semantic_bonus)
            
            # 完全匹配优先记录
            if match_type == '完全匹配':
                exact_match = {
                    'user_field': user_field,
                    'matched_cn': target_cn,
                    'matched_en': target_en,
                    'matched_interface': target_interface,
                    'source': target_interface,
                    'match_type': match_type,
                    'score': total_score
                }
            elif best_match is None or total_score > best_match['score']:
                best_match = {
                    'user_field': user_field,
                    'matched_cn': target_cn,
                    'matched_en': target_en,
                    'matched_interface': target_interface,
                    'source': target_interface,
                    'match_type': match_type,
                    'score': total_score
                }
    
    # 完全匹配优先返回
    if exact_match:
        return exact_match
    
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
        # 检查是否有自定义匹配文件上传
        match_file = request.files.get('match_file')
        if match_file and match_file.filename:
            ext = os.path.splitext(match_file.filename)[1].lower()
            if ext in ['.xlsx', '.xls']:
                match_filepath = os.path.join(app.config['UPLOAD_FOLDER'], 'match_' + match_file.filename)
                match_file.save(match_filepath)
                load_match_data(match_filepath)
        
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
    app.run(host='0.0.0.0', port=5003, debug=True)