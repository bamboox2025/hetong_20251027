#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
超级无敌炫酷数据自动化办公助手 - Vercel 云端部署版
前端：index.html（HTML5 上传）
后端：Flask + 文件上传 + 临时目录
"""

import os
import shutil
import pandas as pd
from flask import Flask, render_template, request, jsonify, send_from_directory
from flask_cors import CORS
from werkzeug.utils import secure_filename

# ==================== 初始化 ====================
app = Flask(__name__, static_folder='static', template_folder='templates')
CORS(app)

# 配置
UPLOAD_FOLDER = '/tmp/uploads'      # Vercel 临时上传目录
OUTPUT_FOLDER = '/tmp/output'       # Vercel 临时输出目录
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024  # 10MB

# 支持的文件类型
ALLOWED_EXTENSIONS = {
    '.doc', '.docx', '.txt', '.pdf', '.rtf',
    '.xls', '.xlsx', '.csv', '.ods',
    '.ppt', '.pptx',
    '.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff'
}

# 全局变量
selected_files = []     # [{"name": "", "path": ""}]
name_list = []          # 名称列表
excel_data = []         # Excel 数据 [[], []]
folder_levels = []      # ["A", "B"]
data_source = 'name'    # 'name' 或 'excel'
output_folder = OUTPUT_FOLDER

# ==================== 工具函数 ====================
def allowed_file(filename):
    return os.path.splitext(filename)[1].lower() in ALLOWED_EXTENSIONS

def column_to_index(column):
    try:
        return ord(column.upper()) - ord('A')
    except:
        return -1

# ==================== 核心处理函数（完整保留您的逻辑） ====================
def process_generation(selected_files, name_list, excel_data, folder_levels, data_source, output_dir):
    result = {
        "success_count": 0,
        "fail_count": 0,
        "folder_count": 0,
        "fail_details": []
    }

    valid_levels = [level for level in folder_levels if level.strip()]
    if not valid_levels:
        result["fail_details"].append("没有有效的文件夹层级设置")
        return result

    if data_source == 'name':
        if not name_list:
            result["fail_details"].append("未获取到名称列表数据")
            return result

        for name_idx, name in enumerate(name_list):
            try:
                folder_parts = []
                for level_idx, level in enumerate(valid_levels):
                    folder_name = name if level_idx == 0 else level if not level.isalpha() else name
                    clean_name = ''.join(c for c in folder_name if c not in '/\\:*?"<>|')
                    if not clean_name:
                        clean_name = f"文件夹_{name_idx}_{level_idx}"
                    folder_parts.append(clean_name)
                
                full_path = os.path.join(output_dir, *folder_parts)
                os.makedirs(full_path, exist_ok=True)
                result["folder_count"] += 1
                
                for file_info in selected_files:
                    try:
                        src = file_info["path"]
                        if not os.path.exists(src):
                            result["fail_details"].append(f"源文件不存在：{file_info['name']}")
                            result["fail_count"] += 1
                            continue
                        if os.path.getsize(src) > 10 * 1024 * 1024:
                            result["fail_details"].append(f"文件过大：{file_info['name']}")
                            result["fail_count"] += 1
                            continue
                        dst = os.path.join(full_path, file_info["name"])
                        shutil.copy2(src, dst)
                        result["success_count"] += 1
                    except Exception as e:
                        result["fail_details"].append(f"复制失败：{str(e)}")
                        result["fail_count"] += 1
            except Exception as e:
                result["fail_details"].append(f"处理第 {name_idx+1} 个名称失败：{str(e)}")
                result["fail_count"] += len(selected_files)

    elif data_source == 'excel':
        if not excel_data:
            result["fail_details"].append("未获取到Excel数据")
            return result

        excel_level_indices = []
        for level in valid_levels:
            idx = column_to_index(level)
            if 0 <= idx < len(excel_data[0]) if excel_data else False:
                excel_level_indices.append(idx)
            else:
                result["fail_details"].append(f"无效列：{level}")

        if not excel_level_indices:
            result["fail_details"].append("没有有效的Excel列")
            return result

        for row_idx, row in enumerate(excel_data):
            try:
                folder_parts = []
                for idx in excel_level_indices:
                    cell = row[idx] if idx < len(row) else ""
                    clean_name = ''.join(c for c in str(cell) if c not in '/\\:*?"<>|')
                    if not clean_name:
                        clean_name = f"文件夹_{row_idx}_{idx}"
                    folder_parts.append(clean_name)
                
                full_path = os.path.join(output_dir, *folder_parts)
                os.makedirs(full_path, exist_ok=True)
                result["folder_count"] += 1
                
                for file_info in selected_files:
                    try:
                        src = file_info["path"]
                        if not os.path.exists(src): continue
                        dst = os.path.join(full_path, file_info["name"])
                        shutil.copy2(src, dst)
                        result["success_count"] += 1
                    except:
                        result["fail_count"] += 1
            except Exception as e:
                result["fail_count"] += len(selected_files)

    return result

# ==================== 路由 ====================

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/static/<path:filename>')
def static_files(filename):
    return send_from_directory('static', filename)

# 上传源文件
@app.route('/upload-source-files', methods=['POST'])
def upload_source_files():
    global selected_files
    selected_files = []
    files = request.files.getlist('files')
    for file in files:
        if file and file.filename and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(path)
            size_mb = os.path.getsize(path) / (1024 * 1024)
            selected_files.append({"name": filename, "path": path, "size": f"{size_mb:.2f} MB"})
    return jsonify({"status": "success", "files": selected_files})

# 上传 name.txt
@app.route('/upload-name-file', methods=['POST'])
def upload_name_file():
    global name_list
    file = request.files.get('file', None)
    if file and file.filename.endswith('.txt'):
        path = os.path.join(app.config['UPLOAD_FOLDER'], 'name.txt')
        file.save(path)
        with open(path, 'r', encoding='utf-8', errors='replace') as f:
            name_list = [line.strip() for line in f if line.strip() and not line.startswith('#')]
        return jsonify({"status": "success", "count": len(name_list), "names": name_list[:10]})
    return jsonify({"status": "error", "message": "无效文件"})

# 上传 Excel
@app.route('/upload-excel-file', methods=['POST'])
def upload_excel_file():
    global excel_data
    file = request.files.get('file', None)
    if file and file.filename.endswith(('.xlsx', '.xls')):
        path = os.path.join(app.config['UPLOAD_FOLDER'], 'data.xlsx')
        file.save(path)
        try:
            df = pd.read_excel(path, engine='openpyxl')
            columns = [chr(65 + i) for i in range(len(df.columns))]
            excel_data = [row.tolist() for _, row in df.iterrows()]
            preview = [row[:10] for row in excel_data[:10]]  # 前10行
            return jsonify({
                "status": "success",
                "rowCount": len(excel_data),
                "columns": columns,
                "data": preview
            })
        except Exception as e:
            return jsonify({"status": "error", "message": str(e)})
    return jsonify({"status": "error", "message": "无效文件"})

# 设置层级
@app.route('/set-folder-levels', methods=['POST'])
def set_folder_levels():
    global folder_levels, data_source
    data = request.json
    folder_levels = data.get('levels', [])
    data_source = data.get('dataSource', 'name')
    return jsonify({"status": "success"})

# 开始处理
@app.route('/process', methods=['POST'])
def process():
    global selected_files, name_list, excel_data, folder_levels, data_source, output_folder
    result = process_generation(
        selected_files, name_list, excel_data,
        folder_levels, data_source, output_folder
    )
    return jsonify({
        "status": "completed",
        "result": result,
        "output_folder": output_folder,
        "file_count": len(selected_files)
    })

# 下载输出（可选）
@app.route('/download/<path:filename>')
def download_file(filename):
    return send_from_directory(output_folder, filename)

# ==================== Vercel 入口 ====================
if __name__ == '__main__':
    app.run(host='127.0.0.1', port=5000, debug=True)
else:
    from waitress import serve
    port = int(os.environ.get('PORT', 5000))
    serve(app, host='0.0.0.0', port=port)