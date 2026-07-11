"""
Lua 配置转换器 —— 主入口（无控制台版本）
将 Excel 配置表转换为 UnLua 可读的 Lua 数据文件。

使用 .pyw 扩展名，Windows 会用 pythonw.exe 启动，不显示控制台窗口。
如需调试，可运行 main.py（带控制台，可看到 print 输出）。
"""
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from ui.main_window import MainWindow


def main():
    app = MainWindow()
    app.run()


if __name__ == "__main__":
    main()
