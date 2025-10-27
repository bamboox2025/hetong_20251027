#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
智能环境配置脚本
自动创建/激活虚拟环境，配置国内源，安装依赖
"""


import os
import sys
import subprocess
import platform
import webbrowser
import time

def get_python_executable():
    """Get path of current Python executable"""
    return sys.executable

def get_venv_path():
    """Get virtual environment path (isolated, no local env pollution)"""
    return os.path.join(os.getcwd(), 'venv')

def is_venv_exists():
    """Check if virtual environment already exists"""
    venv_path = get_venv_path()
    if platform.system() == 'Windows':
        return os.path.exists(os.path.join(venv_path, 'Scripts', 'python.exe'))
    else:
        return os.path.exists(os.path.join(venv_path, 'bin', 'python'))

def create_venv():
    """Create isolated virtual environment"""
    print("Creating virtual environment (isolated, no local env pollution)...")
    python_exe = get_python_executable()
    venv_path = get_venv_path()
    
    try:
        subprocess.check_call([python_exe, '-m', 'venv', venv_path], 
                              stdout=subprocess.DEVNULL, 
                              stderr=subprocess.STDOUT)
        print("Virtual environment created successfully!")
        return True
    except Exception as e:
        print(f"Failed to create virtual environment: {str(e)}")
        return False

def get_venv_python():
    """Get Python executable path inside virtual environment"""
    venv_path = get_venv_path()
    if platform.system() == 'Windows':
        return os.path.join(venv_path, 'Scripts', 'python.exe')
    else:
        return os.path.join(venv_path, 'bin', 'python')

def get_venv_pip():
    """Get pip executable path inside virtual environment"""
    venv_path = get_venv_path()
    if platform.system() == 'Windows':
        return os.path.join(venv_path, 'Scripts', 'pip.exe')
    else:
        return os.path.join(venv_path, 'bin', 'pip')

def configure_pip_mirror():
    """Configure domestic PyPI mirror for fast installation"""
    print("Configuring domestic PyPI mirror (for fast download)...")
    pip_exe = get_venv_pip()
    # Priority domestic mirrors (Tsinghua > Aliyun > USTC)
    mirrors = [
        "https://pypi.tuna.tsinghua.edu.cn/simple",
        "https://mirrors.aliyun.com/pypi/simple/",
        "https://pypi.mirrors.ustc.edu.cn/simple/"
    ]
    
    try:
        # Try Tsinghua mirror first
        result = subprocess.run([pip_exe, 'config', 'set', 'global.index-url', mirrors[0]],
                              capture_output=True, text=True)
        
        if result.returncode == 0:
            print(f"PyPI mirror configured: {mirrors[0]}")
            return True
        else:
            print(f"Tsinghua mirror failed, trying Aliyun...")
            result = subprocess.run([pip_exe, 'config', 'set', 'global.index-url', mirrors[1]],
                                  capture_output=True, text=True)
            
            if result.returncode == 0:
                print(f"PyPI mirror configured: {mirrors[1]}")
                return True
            else:
                print(f"Aliyun mirror failed, trying USTC...")
                result = subprocess.run([pip_exe, 'config', 'set', 'global.index-url', mirrors[2]],
                                      capture_output=True, text=True)
                
                if result.returncode == 0:
                    print(f"PyPI mirror configured: {mirrors[2]}")
                    return True
                else:
                    print("Domestic mirror config failed, use default PyPI source")
                    return False
                    
    except Exception as e:
        print(f"Error configuring PyPI mirror: {str(e)}")
        return False

def install_dependencies():
    """Install dependencies from requirements_minimal.txt (fallback to basic deps if file missing)"""
    print("Installing dependencies (from requirements_minimal.txt)...")
    pip_exe = get_venv_pip()
    req_file = "requirements_minimal.txt"
    requirements = []
    
    # Step 1: Read requirements from file
    try:
        with open(req_file, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                # Skip comments, empty lines
                if not line or line.startswith('#'):
                    continue
                requirements.append(line)
        
        if requirements:
            print(f"Loaded {len(requirements)} dependencies from {req_file}")
        else:
            print(f"No valid dependencies in {req_file}, use fallback list")
            requirements = ['flask>=2.0.0', 'flask-cors>=3.0.0']
    
    # Handle file not found: use fallback dependencies
    except FileNotFoundError:
        print(f"{req_file} not found, use fallback dependencies")
        requirements = ['flask>=2.0.0', 'flask-cors>=3.0.0']
    except Exception as e:
        print(f"Error reading {req_file}: {str(e)}, use fallback dependencies")
        requirements = ['flask>=2.0.0', 'flask-cors>=3.0.0']
    
    # Step 2: Upgrade pip first
    try:
        print("Upgrading pip...")
        subprocess.check_call([pip_exe, 'install', '--upgrade', 'pip'],
                              stdout=subprocess.DEVNULL, 
                              stderr=subprocess.STDOUT)
    except Exception as e:
        print(f"pip upgrade warning: {str(e)} (continue installing dependencies)")
    
    # Step 3: Install dependencies with domestic mirror
    try:
        print("Installing dependencies (with domestic mirror)...")
        subprocess.check_call([
            pip_exe, 'install',
            '-i', 'https://pypi.tuna.tsinghua.edu.cn/simple',  # Force domestic mirror
            '--no-cache-dir',  # Avoid cache issues
            *requirements
        ], stdout=subprocess.DEVNULL, stderr=subprocess.STDOUT)
        
        print("All dependencies installed successfully!")
        return True
    except subprocess.CalledProcessError as e:
        print(f"Dependency install failed: {str(e)}")
        return False
    except Exception as e:
        print(f"Error installing dependencies: {str(e)}")
        return False

def start_application():
    """Start Flask application and open browser"""
    print("Starting application...")
    python_exe = get_venv_python()
    app_file = "app.py"
    
    # Check if app.py exists
    if not os.path.exists(app_file):
        print(f"Error: {app_file} not found in current directory!")
        return False
    
    try:
        # Start Flask app in background
        process = subprocess.Popen([python_exe, app_file],
                                  stdout=subprocess.PIPE,
                                  stderr=subprocess.PIPE,
                                  text=True)
        
        # Wait for server to start (3 seconds)
        print("Waiting for server to start...")
        time.sleep(3)
        
        # Check if process is still running
        if process.poll() is not None:
            stdout, stderr = process.communicate()
            print(f"App startup failed (stdout): {stdout}")
            print(f"App startup failed (stderr): {stderr}")
            return False
        
        # Open browser automatically
        print("Opening browser (http://127.0.0.1:5000)...")
        webbrowser.open('http://127.0.0.1:5000')
        
        print("Application started successfully!")
        print("Access URL: http://127.0.0.1:5000")
        print("Press Ctrl+C to stop the application")
        
        # Keep script running
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            print("\nStopping application...")
            process.terminate()
            process.wait()
            print("Application stopped")
            
        return True
    except Exception as e:
        print(f"Error starting application: {str(e)}")
        return False

def main():
    """Main workflow"""
    print("=" * 60)
    print("        File Batch Tool - Smart Environment Setup")
    print("=" * 60)
    
    # Check Python version (minimum 3.6)
    python_version = sys.version_info
    if python_version < (3, 6):
        print(f"Error: Python 3.6+ is required (current: {python_version.major}.{python_version.minor})")
        sys.exit(1)
    
    print(f"Current Python: {python_version.major}.{python_version.minor}.{python_version.micro}")
    print(f"Current OS: {platform.system()} {platform.release()}")
    print(f"Virtual Env Path: {get_venv_path()}")
    
    # Step 1: Create virtual environment if not exists
    if not is_venv_exists():
        print("Virtual environment not found, creating...")
        if not create_venv():
            print("Failed to create virtual environment, exit")
            sys.exit(1)
    else:
        print("Virtual environment already exists")
    
    # Step 2: Configure PyPI mirror
    configure_pip_mirror()
    
    # Step 3: Install dependencies
    if not install_dependencies():
        print("Dependency install failed, continue? (y/n)")
        choice = input().strip().lower()
        if choice != 'y':
            sys.exit(1)
    
    # Step 4: Start application
    if not start_application():
        print("Application startup failed")
        sys.exit(1)
    
    print("\nSetup completed successfully")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nScript interrupted by user")
    except Exception as e:
        print(f"\nScript error: {str(e)}")
    finally:
        input("\nPress Enter to exit...")