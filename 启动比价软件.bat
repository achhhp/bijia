@echo off
chcp 65001 >nul
cls

echo ========================================
echo 供应商比价软件启动器
echo ========================================
echo 正在检查环境...

:: 检查Python是否安装
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo 错误: 未找到Python环境
    echo 请先下载并安装Python 3.6或更高版本
    echo 下载地址: https://www.python.org/downloads/
    echo 安装时请勾选 "Add Python to PATH" 选项
    pause
    exit /b 1
)

echo Python环境检查通过

:: 创建必要的目录结构
if not exist "templates" (
    echo 创建templates目录...
    md templates
)

:: 创建requirements.txt文件
echo 创建依赖配置文件...
echo Flask==2.0.1>requirements.txt
echo pandas==1.3.3>>requirements.txt
echo openpyxl==3.0.9>>requirements.txt

:: 安装依赖
echo 正在安装依赖包...
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple

if %errorlevel% neq 0 (
    echo 错误: 依赖安装失败
    echo 请检查网络连接后重试
    pause
    exit /b 1
)

echo 依赖安装完成

:: 检查web_app.py文件是否存在
if not exist "web_app.py" (
    echo 错误: 未找到web_app.py文件
    echo 请确保该文件与本批处理文件在同一目录
    pause
    exit /b 1
)

:: 检查index.html文件是否存在
if not exist "templates\index.html" (
    echo 错误: 未找到templates\index.html文件
    echo 请确保该文件存在
    pause
    exit /b 1
)

:: 启动应用
echo 正在启动应用...
echo ========================================
echo 应用将在 http://127.0.0.1:5000 上运行
echo 按 Ctrl+C 停止应用
echo ========================================
python web_app.py

pause