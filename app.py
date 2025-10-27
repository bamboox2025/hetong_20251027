#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
多级文件夹生成工具 - 修复Excel数据处理问题
"""

import os
import shutil
import platform
import tkinter as tk
from tkinter import filedialog
import pandas as pd
from flask import Flask, render_template, request, jsonify
from flask_cors import CORS

# 初始化Flask应用
app = Flask(__name__)
CORS(app)  # 解决跨域问题

# 全局变量存储临时数据
selected_files = []  # 选中的文件路径列表
name_list = []       # name列表（手动输入或从txt读取）
excel_data = []      # Excel数据（修复：确保正确存储）
folder_levels = []   # 文件夹层级设置
data_source = 'name' # 数据来源：name/excel
source_folder = ""   # 源文件夹路径
output_folder = ""   # 输出文件夹路径

# 支持的文件类型（扩展名）
SUPPORTED_EXTENSIONS = {
    '.doc', '.docx', '.txt', '.pdf', '.rtf',
    '.xls', '.xlsx', '.csv', '.ods',
    '.ppt', '.pptx',
    '.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff'
}

def get_default_output_folder():
    """获取默认输出文件夹（桌面）"""
    if platform.system() == 'Windows':
        return os.path.join(os.environ['USERPROFILE'], 'Desktop')
    elif platform.system() == 'Darwin':  # macOS
        return os.path.join(os.environ['HOME'], 'Desktop')
    else:  # Linux
        return os.path.join(os.environ['HOME'], 'Desktop')

def select_folder_dialog():
    """打开文件夹选择对话框"""
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    folder_path = filedialog.askdirectory(title="选择文件夹")
    root.destroy()
    return folder_path if folder_path else ""

def select_file_dialog(filetypes):
    """打开文件选择对话框"""
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    file_path = filedialog.askopenfilename(
        title="选择文件",
        filetypes=filetypes
    )
    root.destroy()
    return file_path if file_path else ""

def get_files_in_folder(folder_path):
    """获取文件夹中所有支持的文件"""
    if not folder_path or not os.path.isdir(folder_path):
        return []
    
    files = []
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        if os.path.isfile(file_path):
            ext = os.path.splitext(filename)[1].lower()
            if ext in SUPPORTED_EXTENSIONS:
                file_size_mb = os.path.getsize(file_path) / (1024 * 1024)
                if file_size_mb <= 10:
                    files.append({
                        "name": filename,
                        "path": file_path,
                        "size": f"{file_size_mb:.2f} MB"
                    })
    return sorted(files, key=lambda x: x["name"])

def read_name_file(file_path):
    """从name.txt读取文件夹名称"""
    if not file_path or not os.path.isfile(file_path):
        return []
    
    try:
        with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
            names = []
            for line in f:
                line = line.strip()
                if line and not line.startswith('#'):
                    names.append(line)
            return names
    except Exception as e:
        print(f"读取name.txt错误：{str(e)}")
        return []

def read_excel_file(file_path):
    """读取Excel文件并返回数据（修复：确保正确解析）"""
    try:
        # 修复：明确指定引擎，确保兼容性
        df = pd.read_excel(file_path, engine='openpyxl')
        
        # 获取列名（A, B, C, ...）
        columns = []
        for i in range(len(df.columns)):
            columns.append(chr(65 + i))
        
        # 转换数据为列表（修复：确保正确处理空值）
        data = []
        for _, row in df.iterrows():
            row_data = []
            for cell in row:
                if pd.isna(cell):
                    row_data.append("")
                else:
                    row_data.append(str(cell).strip())
            data.append(row_data)
        
        return {
            "columns": columns,
            "data": data,
            "rowCount": len(data)
        }
    except Exception as e:
        print(f"读取Excel错误：{str(e)}")
        return None

def column_to_index(column):
    """将列名转换为索引"""
    try:
        return ord(column.upper()) - ord('A')
    except:
        return -1

def process_generation(selected_files, name_list, excel_data, folder_levels, data_source, output_dir):
    """核心处理函数"""
    result = {
        "success_count": 0,
        "fail_count": 0,
        "folder_count": 0,
        "fail_details": []
    }

    # 验证层级设置
    valid_levels = []
    for level in folder_levels:
        valid_levels.append(level)
    
    if not valid_levels:
        result["fail_details"].append("没有有效的文件夹层级设置，无法创建文件夹")
        return result

    # 处理name数据来源
    if data_source == 'name':
        if not name_list:
            result["fail_details"].append("未获取到名称列表数据")
            return result

        for name_idx, name in enumerate(name_list):
            try:
                folder_parts = []
                for level_idx, level in enumerate(valid_levels):
                    if level_idx == 0:
                        folder_name = name
                    else:
                        folder_name = level if not level.isalpha() else name
                    
                    # 清理文件夹名称
                    invalid_chars = '/\\:*?"<>|'
                    clean_name = ''.join([c for c in folder_name if c not in invalid_chars])
                    if not clean_name:
                        clean_name = f"文件夹_{name_idx}_{level_idx}"
                    folder_parts.append(clean_name)
                
                full_folder_path = os.path.join(output_dir, *folder_parts)
                os.makedirs(full_folder_path, exist_ok=True)
                result["folder_count"] += 1
                
                # 复制文件
                for file_path in selected_files:
                    try:
                        if not os.path.exists(file_path):
                            error_msg = f"源文件不存在：{os.path.basename(file_path)}"
                            result["fail_details"].append(error_msg)
                            result["fail_count"] += 1
                            continue

                        file_size_mb = os.path.getsize(file_path) / (1024 * 1024)
                        if file_size_mb > 10:
                            error_msg = f"文件过大（{file_size_mb:.2f} MB）：{os.path.basename(file_path)}"
                            result["fail_details"].append(error_msg)
                            result["fail_count"] += 1
                            continue

                        dest_path = os.path.join(full_folder_path, os.path.basename(file_path))
                        shutil.copy2(file_path, dest_path)
                        result["success_count"] += 1

                    except Exception as e:
                        error_msg = f"复制文件失败：{str(e)}"
                        result["fail_details"].append(error_msg)
                        result["fail_count"] += 1

            except Exception as e:
                error_msg = f"处理第 {name_idx+1} 个名称失败：{str(e)}"
                result["fail_details"].append(error_msg)
                result["fail_count"] += len(selected_files)

    # 处理Excel数据来源（修复：确保正确使用excel_data）
    elif data_source == 'excel':
        if not excel_data or len(excel_data) == 0:
            result["fail_details"].append("未获取到Excel数据或数据为空")
            return result

        # 转换Excel列名到索引
        excel_level_indices = []
        for level in valid_levels:
            idx = column_to_index(level)
            if idx >= 0 and (len(excel_data) > 0 and idx < len(excel_data[0])):
                excel_level_indices.append(idx)
            else:
                result["fail_details"].append(f"无效的列：{level}，已忽略")
        
        if not excel_level_indices:
            result["fail_details"].append("没有有效的Excel列设置，无法创建文件夹")
            return result

        # 处理每一行数据
        for row_idx, row_data in enumerate(excel_data):
            try:
                folder_parts = []
                for level_idx in excel_level_indices:
                    folder_name = row_data[level_idx] if level_idx < len(row_data) else f"未知_{level_idx}"
                    invalid_chars = '/\\:*?"<>|'
                    clean_name = ''.join([c for c in folder_name if c not in invalid_chars])
                    if not clean_name:
                        clean_name = f"文件夹_{row_idx}_{level_idx}"
                    folder_parts.append(clean_name)
                
                full_folder_path = os.path.join(output_dir, *folder_parts)
                os.makedirs(full_folder_path, exist_ok=True)
                result["folder_count"] += 1
                
                # 复制文件
                for file_path in selected_files:
                    try:
                        if not os.path.exists(file_path):
                            error_msg = f"源文件不存在：{os.path.basename(file_path)}"
                            result["fail_details"].append(error_msg)
                            result["fail_count"] += 1
                            continue

                        file_size_mb = os.path.getsize(file_path) / (1024 * 1024)
                        if file_size_mb > 10:
                            error_msg = f"文件过大（{file_size_mb:.2f} MB）：{os.path.basename(file_path)}"
                            result["fail_details"].append(error_msg)
                            result["fail_count"] += 1
                            continue

                        dest_path = os.path.join(full_folder_path, os.path.basename(file_path))
                        shutil.copy2(file_path, dest_path)
                        result["success_count"] += 1

                    except Exception as e:
                        error_msg = f"复制文件失败：{str(e)}"
                        result["fail_details"].append(error_msg)
                        result["fail_count"] += 1

            except Exception as e:
                error_msg = f"处理第 {row_idx+1} 行数据失败：{str(e)}"
                result["fail_details"].append(error_msg)
                result["fail_count"] += len(selected_files)

    return result

# API路由
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/select-source-folder', methods=['POST'])
def select_source_folder():
    global source_folder, selected_files
    source_folder = select_folder_dialog()
    if source_folder:
        files = get_files_in_folder(source_folder)
        selected_files = []
        return jsonify({
            "status": "success",
            "folder": source_folder,
            "files": files
        })
    return jsonify({"status": "error", "message": "未选择文件夹"})

@app.route('/select-files', methods=['POST'])
def select_files():
    global selected_files
    data = request.json
    selected_files = data.get('files', [])
    return jsonify({"status": "success", "count": len(selected_files)})

@app.route('/select-name-file', methods=['POST'])
def select_name_file():
    global name_list
    file_path = select_file_dialog([("文本文件", "*.txt"), ("所有文件", "*.*")])
    if file_path:
        name_list = read_name_file(file_path)
        return jsonify({
            "status": "success",
            "file": file_path,
            "count": len(name_list),
            "names": name_list[:10]
        })
    return jsonify({"status": "error", "message": "未选择文件"})

@app.route('/update-name-list', methods=['POST'])
def update_name_list():
    global name_list
    data = request.json
    name_list = data.get('nameList', [])
    return jsonify({"status": "success", "count": len(name_list)})

@app.route('/select-excel-file', methods=['POST'])
def select_excel_file():
    global excel_data
    file_path = select_file_dialog([("Excel文件", "*.xlsx;*.xls"), ("所有文件", "*.*")])
    if file_path:
        excel_result = read_excel_file(file_path)
        if excel_result:
            # 修复：确保正确赋值给全局变量
            excel_data = excel_result["data"]
            return jsonify({
                "status": "success",
                "file": file_path,
                "rowCount": excel_result["rowCount"],
                "columns": excel_result["columns"],
                "data": excel_result["data"]
            })
        else:
            return jsonify({"status": "error", "message": "读取Excel文件失败"})
    return jsonify({"status": "error", "message": "未选择文件"})

@app.route('/set-folder-levels', methods=['POST'])
def set_folder_levels():
    global folder_levels, data_source
    data = request.json
    folder_levels = data.get('levels', [])
    data_source = data.get('dataSource', 'name')
    return jsonify({"status": "success", "levels": folder_levels, "dataSource": data_source})

@app.route('/set-output-folder', methods=['POST'])
def set_output_folder():
    global output_folder
    data = request.json
    use_default = data.get('useDefault', True)
    main_folder_name = data.get('mainFolder', 'batch_output')
    
    if use_default:
        output_folder = os.path.join(get_default_output_folder(), main_folder_name)
    else:
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        selected_dir = filedialog.askdirectory(title="选择输出位置")
        root.destroy()
        
        if not selected_dir:
            output_folder = os.path.join(get_default_output_folder(), main_folder_name)
        else:
            output_folder = os.path.join(selected_dir, main_folder_name)
    
    os.makedirs(output_folder, exist_ok=True)
    
    return jsonify({
        "status": "success",
        "folder": output_folder
    })

@app.route('/process', methods=['POST'])
def process():
    global selected_files, name_list, excel_data, folder_levels, data_source, output_folder
    
    # 验证必要参数
    if data_source == 'name' and not name_list:
        return jsonify({"status": "error", "message": "未获取到名称列表数据"})
    if data_source == 'excel' and (not excel_data or len(excel_data) == 0):
        return jsonify({"status": "error", "message": "未获取到Excel数据或数据为空"})
    if not folder_levels:
        return jsonify({"status": "error", "message": "未设置文件夹层级"})
    if not output_folder:
        output_folder = os.path.join(get_default_output_folder(), 'batch_output')
        os.makedirs(output_folder, exist_ok=True)
    
    # 执行处理
    result = process_generation(
        selected_files, 
        name_list, 
        excel_data, 
        folder_levels, 
        data_source, 
        output_folder
    )
    return jsonify({
        "status": "completed",
        "result": result,
        "output_folder": output_folder,
        "file_count": len(selected_files)
    })

if __name__ == '__main__':
    if not os.path.exists('templates'):
        os.makedirs('templates')
    app.run(host='127.0.0.1', port=5000, debug=True, use_reloader=False)