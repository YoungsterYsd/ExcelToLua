"""
Lua 配置转换器 —— 主入口
将 Excel 配置表转换为 UnLua 可读的 Lua 数据文件。
"""
import sys
import os

# 确保项目根目录在 path 中
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from ui.main_window import MainWindow


def main():
    app = MainWindow()
    app.run()


if __name__ == "__main__":
    main()
