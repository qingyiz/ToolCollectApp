import sys
import re
import pandas as pd
from PySide6.QtWidgets import (QApplication, QMainWindow, QTextEdit,
                               QPushButton, QVBoxLayout, QHBoxLayout,
                               QWidget, QFileDialog, QLabel, QMessageBox,
                               QTableWidget, QTableWidgetItem, QSplitter)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont, QClipboard, QKeySequence, QShortcut


class InventoryParser(QMainWindow):
    def __init__(self, scale=0.8):
        super().__init__()
        # 支持的单位列表，可根据需要扩展
        self.units = ['个', '只', '条', '斤', '两', '公斤', '克',
                      '包', '板', '箱', '盒', '份', '捆', '束', '根', '块', '瓶', '元', '件']
        # 中文数字映射
        self.chinese_nums = {'半': 0.5, '一': 1, '二': 2, '三': 3, '四': 4, '五': 5,
                             '六': 6, '七': 7, '八': 8, '九': 9, '十': 10,
                             '十一': 11, '十二': 12, '十三': 13, '十四': 14, '十五': 15}
        # 窗口占屏幕的比例，默认80%
        self.scale = scale
        self.init_ui()

    def init_ui(self):
        # 设置窗口标题和大小
        self.setWindowTitle("清单解析与转换工具")
        self.resize(800, 600)

        # 创建主部件和布局
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # 添加标题标签
        title_label = QLabel("清单解析与转换工具")
        title_font = QFont()
        title_font.setPointSize(14)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)

        # 创建分割器，用于分隔输入区和预览区
        splitter = QSplitter(Qt.Vertical)

        # 创建输入区域部件
        input_widget = QWidget()
        input_layout = QVBoxLayout(input_widget)

        # 添加说明标签
        desc_label = QLabel("请输入清单内容（支持格式：1，三黄鸡36斤 或 黑脚鸡1只）：")
        input_layout.addWidget(desc_label)

        # 创建文本编辑框
        self.text_edit = QTextEdit()
        self.text_edit.setFont(QFont("SimHei", 10))
        input_layout.addWidget(self.text_edit)

        # 创建按钮布局
        button_layout = QHBoxLayout()

        # 创建解析按钮
        self.parse_btn = QPushButton("解析清单")
        self.parse_btn.setFont(QFont("SimHei", 10))
        self.parse_btn.clicked.connect(self.parse_and_display)
        button_layout.addWidget(self.parse_btn)

        # 创建生成按钮
        self.generate_btn = QPushButton("生成Excel")
        self.generate_btn.setFont(QFont("SimHei", 10))
        self.generate_btn.clicked.connect(self.generate_excel)
        self.generate_btn.setEnabled(False)  # 初始禁用，解析后启用
        button_layout.addWidget(self.generate_btn)

        # 创建复制按钮
        self.copy_btn = QPushButton("复制表格内容")
        self.copy_btn.setFont(QFont("SimHei", 10))
        self.copy_btn.clicked.connect(self.copy_table_content)
        self.copy_btn.setEnabled(False)  # 初始禁用，解析后启用
        button_layout.addWidget(self.copy_btn)

        input_layout.addLayout(button_layout)
        splitter.addWidget(input_widget)

        # 创建表格预览区域
        preview_widget = QWidget()
        preview_layout = QVBoxLayout(preview_widget)

        preview_label = QLabel("解析结果预览（可选中单元格，使用Ctrl+C复制）：")
        preview_layout.addWidget(preview_label)

        # 创建表格
        self.table_widget = QTableWidget()
        self.table_widget.setFont(QFont("SimHei", 10))
        self.table_widget.setColumnCount(4)
        self.table_widget.setHorizontalHeaderLabels(["序号", "商品名称", "单位", "数量"])
        # 设置表格属性
        self.table_widget.horizontalHeader().setStretchLastSection(True)
        self.table_widget.setEditTriggers(QTableWidget.NoEditTriggers)  # 禁止编辑
        self.table_widget.setSelectionMode(QTableWidget.ExtendedSelection)  # 支持多选
        self.table_widget.setSelectionBehavior(QTableWidget.SelectItems)  # 按单元格选择
        preview_layout.addWidget(self.table_widget)

        # 添加Ctrl+C快捷键支持
        self.add_copy_shortcut()

        splitter.addWidget(preview_widget)

        # 设置分割器的初始大小
        splitter.setSizes([300, 400])

        main_layout.addWidget(splitter)

        # 存储解析后的数据
        self.parsed_data = None

        # 显示状态栏信息
        self.statusBar().showMessage("就绪")

        # 调整窗口大小并居中
        # self.adjust_size_screen()

    def adjust_size_screen(self):
        # 获取屏幕可用区域
        screen_geometry = QApplication.primaryScreen().availableGeometry()
        screen_width = screen_geometry.width()
        screen_height = screen_geometry.height()

        # 计算窗口大小（屏幕的一定比例）
        window_width = int(screen_width * self.scale)
        window_height = int(screen_height * self.scale)

        # 设置窗口大小
        self.resize(window_width, window_height)

        # 窗口居中
        frame_geometry = self.frameGeometry()
        center_point = screen_geometry.center()
        frame_geometry.moveCenter(center_point)
        self.move(frame_geometry.topLeft())

    def add_copy_shortcut(self):
        """添加Ctrl+C快捷键支持表格内容复制"""
        shortcut = QShortcut(QKeySequence("Ctrl+C"), self.table_widget)
        shortcut.activated.connect(self.copy_selected_cells)

    def copy_selected_cells(self):
        """复制选中的单元格内容到剪贴板"""
        selected_items = self.table_widget.selectedItems()
        if not selected_items:
            return

        # 按行和列排序选中的单元格
        selected_items.sort(key=lambda x: (x.row(), x.column()))

        # 获取选中区域的行数和列数
        rows = sorted(set(item.row() for item in selected_items))
        cols = sorted(set(item.column() for item in selected_items))

        # 构建表格数据
        clipboard_text = ""
        for row in rows:
            row_data = []
            for col in cols:
                # 查找当前单元格
                for item in selected_items:
                    if item.row() == row and item.column() == col:
                        row_data.append(item.text())
                        break
            clipboard_text += "\t".join(row_data) + "\n"

        # 去除最后一个换行符
        clipboard_text = clipboard_text.rstrip("\n")

        # 设置剪贴板内容
        clipboard = QApplication.clipboard()
        clipboard.setText(clipboard_text)

        self.statusBar().showMessage(f"已复制 {len(selected_items)} 个单元格内容")

    def parse_inventory(self, text):
        """解析输入的清单文本，同时支持带序号多行项目和逗号分隔单行项目"""
        items = []

        # 首先按换行分割成块
        line_blocks = text.split('\n')

        # 处理每个块，再按逗号分割成单个项目
        all_items = []
        for block in line_blocks:
            block = block.strip()
            if not block:
                continue

            # 按中文逗号和英文逗号分割项目，但保留带序号项目的完整性
            if re.match(r'^\d+\s*[，,。.:、]', block):
                all_items.append(block)
            else:
                parts = re.split(r'[，,。.]', block)
                for part in parts:
                    part = part.strip()
                    if part:
                        all_items.append(part)

        # 解析每个项目
        for item_text in all_items:
            try:
                index = len(items) + 1  # 默认序号为当前条目数+1
                rest = item_text

                # 统一处理带序号的项目，兼容多种分隔符：中文逗号、中英文句号、顿号等
                # 匹配模式：数字 + 可能的空白 + 分隔符(,，.:、) + 可能的空白
                # 增强版：支持更多种标点符号作为分隔符
                seq_pattern = r'^(\d+)\s*[,，.。:：、]\s*(.*)$'
                seq_match = re.match(seq_pattern, item_text)

                if seq_match:
                    index = int(seq_match.group(1))
                    rest = seq_match.group(2).strip()

                # 提取数量和单位
                quantity = ""
                unit = ""
                product_name = rest

                # 构建单位正则表达式
                units_regex = '|'.join([re.escape(u) for u in self.units])

                # 标记是否是中文数量匹配
                is_chinese_num = False

                # 匹配数字+单位模式（如1.5斤、2个）
                pattern = rf'(\d+\.?\d*)\s*({units_regex})$'
                match = re.search(pattern, rest)

                if not match:
                    # 匹配中文数量+单位模式（如一只、两个）
                    chinese_nums_regex = '|'.join([re.escape(k) for k in self.chinese_nums.keys()])
                    pattern = rf'({chinese_nums_regex})\s*({units_regex})$'
                    match = re.search(pattern, rest)
                    if match:
                        is_chinese_num = True  # 标记为中文数量匹配
                        num_char = match.group(1)
                        quantity = self.chinese_nums.get(num_char, "")
                        unit = match.group(2)
                        product_name = rest[:match.start()].strip()

                # 只对数字模式进行float转换，中文模式已提前处理
                if match and not is_chinese_num:
                    # 处理数字+单位的匹配结果
                    quantity = float(match.group(1))
                    unit = match.group(2)
                    product_name = rest[:match.start()].strip()

                items.append({
                    '序号': index,
                    '商品名称': product_name,
                    '单位': unit,
                    '数量': quantity
                })

            except Exception as e:
                QMessageBox.warning(self, "解析错误", f"解析项目 '{item_text}' 时出错: {str(e)}")
                continue

        return items

    def parse_and_display(self):
        """解析清单并显示在表格中"""
        # 获取文本编辑框中的内容
        text = self.text_edit.toPlainText()

        if not text:
            QMessageBox.warning(self, "警告", "请输入清单内容")
            return

        # 解析清单
        self.statusBar().showMessage("正在解析清单...")
        self.parsed_data = self.parse_inventory(text)

        if not self.parsed_data:
            QMessageBox.warning(self, "警告", "未能解析任何清单项目")
            self.statusBar().showMessage("就绪")
            return

        # 显示在表格中
        self.table_widget.setRowCount(len(self.parsed_data))

        for row, item in enumerate(self.parsed_data):
            # 序号
            index_item = QTableWidgetItem(str(item['序号']))
            index_item.setTextAlignment(Qt.AlignCenter)
            self.table_widget.setItem(row, 0, index_item)

            # 商品名称
            name_item = QTableWidgetItem(item['商品名称'])
            name_item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            self.table_widget.setItem(row, 1, name_item)

            # 单位
            unit_item = QTableWidgetItem(item['单位'])
            unit_item.setTextAlignment(Qt.AlignCenter)
            self.table_widget.setItem(row, 2, unit_item)

            # 数量
            quantity_item = QTableWidgetItem(str(item['数量']))
            quantity_item.setTextAlignment(Qt.AlignCenter)
            self.table_widget.setItem(row, 3, quantity_item)

        # 调整列宽
        self.table_widget.resizeColumnsToContents()

        # 启用按钮
        self.generate_btn.setEnabled(True)
        self.copy_btn.setEnabled(True)

        self.statusBar().showMessage(f"已解析 {len(self.parsed_data)} 条记录")

    def copy_table_content(self):
        """复制整个表格内容到剪贴板"""
        if not self.parsed_data:
            return

        # 构建制表符分隔的文本
        clipboard_text = "序号\t商品名称\t单位\t数量\n"  # 表头

        for item in self.parsed_data:
            line = f"{item['序号']}\t{item['商品名称']}\t{item['单位']}\t{item['数量']}\n"
            clipboard_text += line

        # 获取剪贴板并设置内容
        clipboard = QApplication.clipboard()
        clipboard.setText(clipboard_text)

        self.statusBar().showMessage("表格内容已复制到剪贴板，可以粘贴到Excel或其他应用中")

    def generate_excel(self):
        """生成Excel文件"""
        if not self.parsed_data:
            QMessageBox.warning(self, "警告", "请先解析清单内容")
            return

        # 创建DataFrame
        df = pd.DataFrame(self.parsed_data)

        # 让用户选择保存位置
        file_path, _ = QFileDialog.getSaveFileName(
            self, "保存Excel文件", "", "Excel Files (*.xlsx);;All Files (*)"
        )

        if not file_path:
            self.statusBar().showMessage("保存已取消")
            return

        # 确保文件以.xlsx结尾
        if not file_path.endswith('.xlsx'):
            file_path += '.xlsx'

        try:
            # 保存为Excel文件
            self.statusBar().showMessage("正在生成Excel文件...")
            df.to_excel(file_path, index=False)
            self.statusBar().showMessage(f"文件已保存至: {file_path}")
            QMessageBox.information(self, "成功", f"Excel文件已成功生成:\n{file_path}")
        except Exception as e:
            self.statusBar().showMessage("生成文件失败")
            QMessageBox.critical(self, "错误", f"生成Excel文件时出错: {str(e)}")


if __name__ == "__main__":
    # 确保中文显示正常
    app = QApplication(sys.argv)

    # 设置全局字体，确保中文显示
    font = QFont("SimHei")
    app.setFont(font)

    window = InventoryParser()
    window.show()
    sys.exit(app.exec_())
