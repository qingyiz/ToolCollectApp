import sys
from PySide6.QtWidgets import QApplication
from toolcollectionapp import StyleManager, ToolCollectionApp

# ------------------------------
# 程序入口
# ------------------------------
if __name__ == "__main__":
    app = QApplication(sys.argv)

    # 初始化并应用全局样式
    style_manager = StyleManager()
    style_manager.apply_global_styles(app)

    # 启动主程序
    window = ToolCollectionApp()
    window.show()
    sys.exit(app.exec())