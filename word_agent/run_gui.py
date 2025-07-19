#!/usr/bin/env python3
"""
Word Agent GUI Launcher
启动Word Agent的图形界面
"""

import sys
import os
import subprocess

def check_dependencies():
    """检查依赖项"""
    try:
        import tkinter
        import docx
        import openai
        print("✓ 所有依赖项已安装")
        return True
    except ImportError as e:
        print(f"✗ 缺少依赖项: {e}")
        print("请运行: pip install -r requirements_gui.txt")
        return False

def main():
    """主函数"""
    print("Word Agent GUI Launcher")
    print("=" * 50)
    
    # 检查依赖项
    if not check_dependencies():
        input("按回车键退出...")
        return
    
    # 启动GUI
    try:
        print("正在启动GUI...")
        from word_agent_gui import main as gui_main
        gui_main()
    except Exception as e:
        print(f"启动GUI失败: {e}")
        input("按回车键退出...")

if __name__ == "__main__":
    main() 