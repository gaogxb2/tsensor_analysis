#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
使用 PyInstaller 将 GUI 程序打包成 exe 文件

使用方法：
    python build_exe.py

注意：建议在全新的虚拟环境中打包，避免打包不必要的依赖
    1. 创建虚拟环境: python -m venv venv_build
    2. 激活虚拟环境
    3. 安装必要依赖: pip install pyinstaller openpyxl
    4. 运行打包: python build_exe.py
"""

import subprocess
import sys
from pathlib import Path


def main():
    """执行打包"""
    # 获取脚本所在目录
    script_dir = Path(__file__).parent
    gui_file = script_dir / "process_temperature_gui.py"
    
    if not gui_file.exists():
        print(f"错误：找不到文件 {gui_file}")
        sys.exit(1)
    
    print("开始打包 GUI 程序...")
    print(f"入口文件: {gui_file}")
    print("提示：建议在全新的虚拟环境中打包，避免包含不必要的依赖包")
    print("-" * 60)
    
    # PyInstaller 命令参数
    cmd = [
        "pyinstaller",
        "--onefile",                    # 打包成单个 exe 文件
        "--windowed",                    # Windows 下隐藏控制台窗口（macOS/Linux 使用 --noconsole）
        "--name=Tsensor温度处理工具",    # 输出 exe 名称
        "--hidden-import=openpyxl",      # 显式导入 openpyxl（确保被包含）
        "--hidden-import=process_temperature_data",  # 显式导入处理模块
        "--clean",                       # 清理临时文件
        # 排除不必要的模块，减小打包体积
        "--exclude-module=matplotlib",   # 排除 matplotlib（如果不需要绘图）
        "--exclude-module=numpy",        # 排除 numpy（如果 openpyxl 不需要）
        "--exclude-module=pandas",       # 排除 pandas
        "--exclude-module=scipy",        # 排除 scipy
        "--exclude-module=IPython",       # 排除 IPython
        "--exclude-module=jupyter",       # 排除 jupyter
        "--exclude-module=notebook",      # 排除 notebook
        "--exclude-module=PyQt5",        # 排除 PyQt5
        "--exclude-module=PyQt6",        # 排除 PyQt6
        "--exclude-module=PySide2",      # 排除 PySide2
        "--exclude-module=PySide6",      # 排除 PySide6
        "--exclude-module=tkinter.test", # 排除 tkinter 测试模块
        str(gui_file),
    ]
    
    # 检测操作系统，调整参数
    if sys.platform == "darwin":  # macOS
        cmd[2] = "--noconsole"  # macOS 使用 --noconsole
    elif sys.platform.startswith("linux"):  # Linux
        cmd[2] = "--noconsole"
    # Windows 使用 --windowed
    
    print(f"执行命令: {' '.join(cmd)}")
    print("-" * 60)
    
    try:
        # 执行打包
        result = subprocess.run(cmd, check=True, cwd=str(script_dir))
        print("-" * 60)
        print("打包完成！")
        print(f"输出目录: {script_dir / 'dist'}")
        if sys.platform == "win32":
            print(f"exe 文件: {script_dir / 'dist' / 'Tsensor温度处理工具.exe'}")
        else:
            print(f"可执行文件: {script_dir / 'dist' / 'Tsensor温度处理工具'}")
    except subprocess.CalledProcessError as e:
        print(f"打包失败: {e}")
        sys.exit(1)
    except FileNotFoundError:
        print("错误：未找到 pyinstaller 命令")
        print("请先安装 PyInstaller: pip install pyinstaller")
        print("\n建议在全新虚拟环境中打包：")
        print("  1. python -m venv venv_build")
        print("  2. 激活虚拟环境（Windows: venv_build\\Scripts\\activate，macOS/Linux: source venv_build/bin/activate）")
        print("  3. pip install pyinstaller openpyxl")
        print("  4. python build_exe.py")
        sys.exit(1)


if __name__ == "__main__":
    main()
