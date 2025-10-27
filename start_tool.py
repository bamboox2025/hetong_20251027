#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文件批量处理工具 - 启动脚本
"""

import os
import sys
import subprocess
import time
import webbrowser
from tkinter import Tk, messagebox

def check_requirements():
    """检查依赖包"""
    try:
        import flask
        import flask_cors
        return True
    except ImportError:
        return False

def install_requirements():
    """安装依赖包"""
    try:
        subprocess.check_call([
            sys.executable, '-m', 'pip', 'install', 
            '--user', 'flask', 'flask-cors'
        ])
        return True
    except Exception as e:
        print(f"安装依赖包失败: {e}")
        return False

def main():
    """主函数"""
    # 检查并安装依赖
    if not check_requirements():
        print("正在安装必要的依赖包...")
        if not install_requirements():
            # 创建Tkinter消息框显示错误
            root = Tk()
            root.withdraw()
            messagebox.showerror(
                "错误", 
                "安装依赖包失败，请手动安装：\npip install flask flask-cors"
            )
            root.destroy()
            sys.exit(1)
    
    # 获取当前目录
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 启动Flask应用
    try:
        # 使用subprocess启动，避免阻塞
        subprocess.Popen([
            sys.executable, os.path.join(current_dir, 'app.py')
        ])
        
        # 等待服务启动
        time.sleep(2)
        
        # 打开浏览器
        webbrowser.open('http://127.0.0.1:5000')
        
        print("文件批量处理工具已启动！")
        print("浏览器将自动打开工具界面...")
        print("如果浏览器没有自动打开，请访问: http://127.0.0.1:5000")
        
    except Exception as e:
        print(f"启动应用失败: {e}")
        # 创建Tkinter消息框显示错误
        root = Tk()
        root.withdraw()
        messagebox.showerror(
            "错误", 
            f"启动应用失败: {str(e)}"
        )
        root.destroy()
        sys.exit(1)

if __name__ == '__main__':
    main()