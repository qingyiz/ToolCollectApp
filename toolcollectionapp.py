import sys
import os
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                               QHBoxLayout, QPushButton, QLabel, QGridLayout,
                               QScrollArea)
from PySide6.QtGui import QIcon, QPixmap, QFont, QColor, QPainter
from PySide6.QtCore import Qt, QSize

# 为不同工具设置专属emoji图标
icon_map = {
    "清单解析工具": "📋",
    "Excel提取工具": "📊",
    "Excel合并工具": "🧩",
    "默认工具": "🔧"
}


# ------------------------------
# 全局样式管理器（简约风格）
# ------------------------------
class StyleManager:
    _instance = None

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance.init_styles()
        return cls._instance

    def init_styles(self):
        """初始化简约风格样式配置"""
        # 基础配色方案（更浅的色调）
        self.colors = {
            "main": "#4a90e2",  # 主色调（浅蓝）
            "secondary": "#6b7c93",  # 次要色（浅灰蓝）
            "light_bg": "#f7f9fc",  # 浅色背景
            "hover": "#ebf3fc",  # 悬浮色
            "pressed": "#dceafd",  # 点击色
            "text": "#32325d",  # 文本色
            "text_light": "#8898aa",  # 浅色文本
            "border": "#e2e8f0"  # 边框色
        }

        # 字体配置
        self.fonts = {
            "family": "Microsoft YaHei",
            "title_size": 18,
            "subtitle_size": 14,
            "normal_size": 12,
            "small_size": 10
        }

        # 布局参数
        self.layout = {
            "border_radius": 6,  # 更小的圆角，更简约
            "spacing": 10,
            "margin": 12
        }

        # 生成全局样式表
        self.global_style = self._generate_global_style()
        self.tool_button_style = self._generate_tool_button_style()

    def _generate_global_style(self):
        """生成全局应用的样式表"""
        c = self.colors
        f = self.fonts
        l = self.layout

        return f"""
            /* 全局字体设置 */
            * {{
                font-family: {f['family']};
                font-size: {f['normal_size']}px;
            }}

            /* 主窗口样式 */
            QMainWindow, QWidget {{
                background-color: {c['light_bg']};
                color: {c['text']};
            }}

            /* 标题标签 */
            QLabel#title_label {{
                color: {c['main']};
                font-size: {f['title_size']}px;
                font-weight: bold;
                margin-bottom: {l['spacing']}px;
            }}

            /* 描述标签 */
            QLabel#desc_label {{
                color: {c['text_light']};
                font-size: {f['small_size']}px;
            }}

            /* 分组框 */
            QGroupBox {{
                border: 1px solid {c['border']};
                border-radius: {l['border_radius']}px;
                margin-top: 10px;
                font-size: {f['subtitle_size']}px;
                color: {c['secondary']};
                font-weight: bold;
            }}

            QGroupBox::title {{
                subcontrol-origin: margin;
                left: 15px;
                padding: 0 5px;
            }}

            /* 按钮样式 */
            QPushButton {{
                background-color: {c['main']};
                color: white;
                border: none;
                border-radius: {l['border_radius']}px;
                padding: 6px 14px;
                font-weight: 500;
                transition: all 0.2s ease;
            }}

            QPushButton:hover {{
                background-color: #5a9dec;
                transform: translateY(-1px);
            }}

            QPushButton:pressed {{
                background-color: #3d85d6;
                transform: translateY(0);
            }}

            QPushButton:disabled {{
                background-color: {c['text_light']};
                color: #cbd5e1;
                transform: none;
            }}

            /* 状态栏 */
            QStatusBar {{
                background-color: white;
                color: {c['text_light']};
                font-size: {f['small_size']}px;
                border-top: 1px solid {c['border']};
            }}

            /* 滚动区域 */
            QScrollArea {{
                border: none;
            }}
        """

    def _generate_tool_button_style(self):
        """生成工具按钮专用样式表（优化图标颜色和谐性）"""
        c = self.colors
        f = self.fonts
        l = self.layout

        return f"""
            ToolButton {{
                background-color: white;
                border: 1px solid {c['border']};
                border-radius: {l['border_radius']}px;
                min-width: 200px;
                min-height: 160px;
                max-width: 200px;
                max-height: 160px;
                transition: all 0.2s ease;
            }}

            ToolButton:hover {{
                background-color: {c['hover']};
                border-color: {c['main']};
                transform: translateY(-2px);
            }}

            ToolButton:pressed {{
                background-color: {c['pressed']};
                border-color: {c['main']};
                transform: translateY(0);
            }}

            /* 工具按钮中的元素样式 */
            ToolButton QLabel#tool_name {{
                background-color: transparent;
                color: {c['text']};
                font-size: {f['subtitle_size']}px;
                font-weight: bold;
                margin-top: 5px;
            }}

            ToolButton QLabel#tool_desc {{
                background-color: transparent;
                color: {c['text_light']};
                font-size: {f['small_size']}px;
                margin-top: 5px;
                word-wrap: true;
            }}

            ToolButton QLabel#tool_icon {{
                background-color: transparent;
                font-size: 32px;
                color: {c['secondary']}; /* 默认图标颜色 */
            }}

            /* 悬停时图标颜色变化，保持和谐 */
            ToolButton:hover QLabel#tool_icon {{
                color: {c['main']}; /* 悬停时与边框颜色一致 */
            }}
        """

    def apply_global_styles(self, app):
        """将全局样式应用到应用程序"""
        app.setStyleSheet(self.global_style)

        # 设置全局字体
        font = QFont(self.fonts['family'], self.fonts['normal_size'])
        app.setFont(font)


# ------------------------------
# 自定义工具按钮组件
# ------------------------------
class ToolButton(QPushButton):
    def __init__(self, tool_info, parent=None):
        super().__init__(parent)
        self.tool_info = tool_info
        self.style_manager = StyleManager()
        self.init_ui()

    def init_ui(self):
        # 设置按钮布局
        layout = QVBoxLayout(self)
        layout.setContentsMargins(
            self.style_manager.layout['margin'],
            self.style_manager.layout['margin'],
            self.style_manager.layout['margin'],
            self.style_manager.layout['margin']
        )
        layout.setSpacing(8)

        # 工具图标
        self.icon_label = QLabel()
        self.icon_label.setObjectName("tool_icon")
        self.icon_label.setAlignment(Qt.AlignCenter)
        self._set_icon()
        layout.addWidget(self.icon_label)

        # 工具名称
        self.name_label = QLabel(self.tool_info["name"])
        self.name_label.setObjectName("tool_name")
        self.name_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.name_label)

        # 工具描述
        self.desc_label = QLabel(self.tool_info["description"])
        self.desc_label.setObjectName("tool_desc")
        self.desc_label.setAlignment(Qt.AlignCenter)
        self.desc_label.setWordWrap(True)
        layout.addWidget(self.desc_label)

        # 应用工具按钮样式
        self.setStyleSheet(self.style_manager.tool_button_style)

        # 绑定点击事件
        self.clicked.connect(self.tool_info["launch_func"])

    def _set_icon(self):
        """设置工具图标"""
        tool = self.tool_info
        icon_path = tool["icon"]

        if icon_path and os.path.exists(icon_path):
            pixmap = QPixmap(icon_path).scaled(56, 56, Qt.KeepAspectRatio,
                                               Qt.SmoothTransformation)
            self.icon_label.setPixmap(pixmap)
        else:
            self.icon_label.setText(icon_map.get(tool["name"], "🔧"))
            icon_font = QFont()
            icon_font.setPointSize(28)
            self.icon_label.setFont(icon_font)


# ------------------------------
# 主工具合集界面
# ------------------------------
class ToolCollectionApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.tools = []
        self.style_manager = StyleManager()  # 初始化样式管理器
        self.init_ui()
        self.register_tools()

    def init_ui(self):
        # 窗口基础设置
        self.setWindowTitle("工具合集")
        self.setGeometry(100, 100, 850, 550)

        # 中心部件与主布局
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)

        # 标题区域
        title_label = QLabel("工具合集")
        title_label.setObjectName("title_label")  # 仅设置对象名
        main_layout.addWidget(title_label)

        # 说明文本
        desc_label = QLabel("点击下方工具按钮启动对应功能")
        desc_label.setObjectName("desc_label")  # 仅设置对象名
        main_layout.addWidget(desc_label)

        # 工具区域（滚动布局）
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)

        tools_container = QWidget()
        self.tools_layout = QGridLayout(tools_container)
        self.tools_layout.setContentsMargins(10, 10, 10, 10)
        self.tools_layout.setSpacing(15)  # 工具按钮间距

        scroll_area.setWidget(tools_container)
        main_layout.addWidget(scroll_area)

        # 状态栏
        self.statusBar().showMessage("就绪")

        self.adjust_size_screen()

    def adjust_size_screen(self, widget=None, scale=0.6):
        if widget is None:
            widget = self
        # 获取屏幕可用区域
        screen_geometry = QApplication.primaryScreen().availableGeometry()
        screen_width = screen_geometry.width()
        screen_height = screen_geometry.height()

        # 计算窗口大小（屏幕的一定比例）
        window_width = int(screen_width * scale)
        window_height = int(screen_height * scale)

        # 设置窗口大小
        widget.resize(window_width, window_height)

        # 窗口居中
        frame_geometry = widget.frameGeometry()
        center_point = screen_geometry.center()
        frame_geometry.moveCenter(center_point)
        widget.move(frame_geometry.topLeft())

    def create_emoji_icon(self, emoji, size=64):
        """将emoji转换为QIcon"""
        # 创建透明像素图
        pixmap = QPixmap(size, size)
        pixmap.fill(Qt.transparent)

        # 在像素图上绘制emoji
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.Antialiasing)

        # 设置字体和大小
        font = painter.font()
        font.setPointSize(size // 2)  # emoji大小为图标一半
        painter.setFont(font)

        # 居中绘制emoji
        painter.drawText(pixmap.rect(), Qt.AlignCenter, emoji)
        painter.end()

        # 转换为QIcon并返回
        return QIcon(pixmap)

    def register_tools(self):
        """注册所有工具"""
        self.tools = [
            {
                "name": "清单解析工具",
                "description": "将文本清单解析为表格并导出Excel",
                "icon": self.get_icon_path("清单解析工具"),
                "launch_func": self.launch_inventory_parser
            },
            {
                "name": "Excel合并工具",
                "description": "合并execl表格",
                "icon": self.get_icon_path("Excel合并工具"),
                "launch_func": self.launch_merge_execl
            },
            {
                "name": "Excel提取工具",
                "description": "批量提取Excel文件中的合计数据",
                "icon": self.get_icon_path("Excel提取工具"),
                "launch_func": self.launch_excel_processor
            },
            # 新增工具只需在这里添加
            # {
            #     "name": "新工具名称",
            #     "description": "新工具的功能描述",
            #     "icon": self.get_icon_path("new_icon.png"),
            #     "launch_func": self.launch_new_tool
            # }
        ]

        # 添加工具按钮到界面（每行3个）
        tools_per_row = 3
        for i, tool in enumerate(self.tools):
            tool_btn = ToolButton(tool)
            row = i // tools_per_row
            col = i % tools_per_row
            self.tools_layout.addWidget(tool_btn, row, col)

    def get_icon_path(self, icon_name):
        """获取图标路径"""
        icon_dir = "icons"
        if not os.path.exists(icon_dir):
            # os.makedirs(icon_dir)
            return ""
        return os.path.join(icon_dir, icon_name)

    # 工具启动函数
    def launch_inventory_parser(self):
        """启动清单解析工具"""
        # 实际使用时替换为您的工具类
        from inventory_to_excel import InventoryParser
        self.inventory_window = InventoryParser()
        self.inventory_window.setWindowIcon(self.create_emoji_icon(icon_map.get("清单解析工具", "🔧")))
        self.adjust_size_screen(self.inventory_window, 0.5)
        self.inventory_window.show()
        self.statusBar().showMessage("已启动：清单解析工具")

    def launch_excel_processor(self):
        """启动Excel提取工具"""
        # 实际使用时替换为您的工具类
        from total import ExcelProcessor
        self.excel_window = ExcelProcessor()
        self.excel_window.setWindowIcon(self.create_emoji_icon(icon_map.get("Excel提取工具", "🔧")))
        self.adjust_size_screen(self.excel_window, 0.5)
        self.excel_window.show()
        self.statusBar().showMessage("已启动：Excel提取工具")

    def launch_merge_execl(self):
        """启动Excel提取工具"""
        # 实际使用时替换为您的工具类
        from merge import ExcelMergerApp
        self.merge_execl = ExcelMergerApp()
        self.merge_execl.setWindowIcon(self.create_emoji_icon(icon_map.get("Excel合并工具", "🔧")))
        self.adjust_size_screen(self.merge_execl, 0.5)
        self.merge_execl.show()
        self.statusBar().showMessage("已启动：Excel合并工具")

    # 新增工具启动函数示例
    # def launch_new_tool(self):
    #     from new_tool import NewToolClass
    #     self.new_tool_window = NewToolClass()
    #     self.new_tool_window.show()
    #     self.statusBar().showMessage("已启动：新工具名称")

