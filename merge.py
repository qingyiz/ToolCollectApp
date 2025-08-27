import sys
import pandas as pd
import os
from PySide6 import QtWidgets, QtGui, QtCore
from PySide6.QtGui import QKeySequence, QShortcut
from PySide6.QtWidgets import (
    QTableWidgetItem,
    QAbstractItemView,
    QMainWindow,
    QFileDialog,
    QDialog,
    QTextEdit,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton
)
from PySide6.QtCore import Qt, Signal


class LogDialog(QDialog):
    """日志弹窗对话框"""

    def __init__(self, parent=None, log_text=""):
        super().__init__(parent)
        self.setWindowTitle("操作日志")
        self.resize(800, 500)
        self.setMinimumSize(600, 400)

        # 主布局
        main_layout = QVBoxLayout(self)

        # 日志文本框
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setStyleSheet("background-color: #f9f9f9;")
        self.log_text.setText(log_text)
        main_layout.addWidget(self.log_text)

        # 按钮布局
        btn_layout = QHBoxLayout()
        btn_layout.addStretch(1)

        # 清空日志按钮
        self.clear_btn = QPushButton("清空日志")
        self.clear_btn.clicked.connect(self.clear_log)
        btn_layout.addWidget(self.clear_btn)

        # 复制日志按钮
        self.copy_btn = QPushButton("复制全部日志")
        self.copy_btn.clicked.connect(self.copy_log)
        btn_layout.addWidget(self.copy_btn)

        main_layout.addLayout(btn_layout)

    def clear_log(self):
        """清空日志内容"""
        self.log_text.clear()
        if self.parent():
            self.parent().log_buffer = ""  # 同步清空主窗口的日志缓存

    def copy_log(self):
        """复制全部日志到剪贴板"""
        log_text = self.log_text.toPlainText()
        if log_text:
            clipboard = QtWidgets.QApplication.clipboard()
            clipboard.setText(log_text)
            self.parent().status_bar.showMessage("✅ 日志已复制到剪贴板", 3000)


class ExcelMergerApp(QMainWindow):
    """Excel数据合并工具主窗口"""
    update_table_signal = Signal(pd.DataFrame)
    log_update_signal = Signal(str)  # 用于更新日志的信号

    def __init__(self):
        super().__init__()
        self.current_merged_df = None  # 存储当前合并后的DataFrame
        self.log_buffer = ""  # 日志缓存，用于保存完整日志
        self.init_ui()
        self.bind_signals()

    def init_ui(self):
        """初始化用户界面"""
        self.setWindowTitle("Excel数据合并工具")
        self.resize(1300, 700)

        # 创建中心部件和主布局
        central_widget = QtWidgets.QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QtWidgets.QVBoxLayout(central_widget)
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(10)

        # ---------------------- 文件选择区域 ----------------------
        file_layout = QtWidgets.QHBoxLayout()
        file_layout.setSpacing(10)

        # 目标文件选择
        target_file_label = QtWidgets.QLabel("目标文件:")
        self.first_file_input = QtWidgets.QLineEdit()
        self.first_file_input.setPlaceholderText("请选择需要填入数据的Excel文件")
        self.first_file_button = QtWidgets.QPushButton("浏览")

        # 数据文件选择
        data_file_label = QtWidgets.QLabel("数据文件:")
        self.second_file_input = QtWidgets.QLineEdit()
        self.second_file_input.setPlaceholderText("请选择提供数据的Excel文件")
        self.second_file_button = QtWidgets.QPushButton("浏览")

        # 添加到文件布局
        file_layout.addWidget(target_file_label)
        file_layout.addWidget(self.first_file_input, 1)
        file_layout.addWidget(self.first_file_button)
        file_layout.addSpacing(20)
        file_layout.addWidget(data_file_label)
        file_layout.addWidget(self.second_file_input, 1)
        file_layout.addWidget(self.second_file_button)
        main_layout.addLayout(file_layout)

        # ---------------------- 参数设置区域 ----------------------
        self.parameters_label = QtWidgets.QLabel("参数设置")
        self.parameters_label.setStyleSheet("font-weight: bold; font-size: 14px;")
        main_layout.addWidget(self.parameters_label)

        param_layout = QtWidgets.QGridLayout()
        param_layout.setSpacing(8)
        param_layout.setColumnStretch(1, 1)
        param_layout.setColumnStretch(3, 1)

        # 目标文件参数（左半部分）
        param_layout.addWidget(QtWidgets.QLabel("目标文件列名行号:"), 0, 0, Qt.AlignRight)
        self.first_row_input = QtWidgets.QSpinBox()
        self.first_row_input.setValue(3)
        self.first_row_input.setRange(1, 100)
        param_layout.addWidget(self.first_row_input, 0, 1)

        param_layout.addWidget(QtWidgets.QLabel("目标文件匹配列名:"), 1, 0, Qt.AlignRight)
        self.first_col_name_input = QtWidgets.QLineEdit()
        self.first_col_name_input.setText("商品名称")
        param_layout.addWidget(self.first_col_name_input, 1, 1)

        param_layout.addWidget(QtWidgets.QLabel("目标文件待填列名:"), 2, 0, Qt.AlignRight)
        self.first_data_col_input = QtWidgets.QLineEdit()
        self.first_data_col_input.setText("价格")
        param_layout.addWidget(self.first_data_col_input, 2, 1)

        # 数据文件参数（右半部分）
        param_layout.addWidget(QtWidgets.QLabel("数据文件列名行号:"), 0, 2, Qt.AlignRight)
        self.second_row_input = QtWidgets.QSpinBox()
        self.second_row_input.setValue(1)
        self.second_row_input.setRange(1, 100)
        param_layout.addWidget(self.second_row_input, 0, 3)

        param_layout.addWidget(QtWidgets.QLabel("数据文件匹配列名:"), 1, 2, Qt.AlignRight)
        self.second_col_name_input = QtWidgets.QLineEdit()
        self.second_col_name_input.setText("商品名称")
        param_layout.addWidget(self.second_col_name_input, 1, 3)

        param_layout.addWidget(QtWidgets.QLabel("数据文件来源列名:"), 2, 2, Qt.AlignRight)
        self.second_data_col_input = QtWidgets.QLineEdit()
        self.second_data_col_input.setText("TG")
        param_layout.addWidget(self.second_data_col_input, 2, 3)

        main_layout.addLayout(param_layout)

        # ---------------------- 功能按钮布局 ----------------------
        btn_main_layout = QtWidgets.QHBoxLayout()
        btn_main_layout.setSpacing(15)

        # 1. 执行合并按钮
        self.run_button = QtWidgets.QPushButton("执行合并")
        self.run_button.setCursor(Qt.PointingHandCursor)
        self.run_button.setFixedWidth(120)

        # 2. 复制全部按钮
        self.copy_all_btn = QtWidgets.QPushButton("复制全部表格内容")
        self.copy_all_btn.setCursor(Qt.PointingHandCursor)
        self.copy_all_btn.setToolTip("复制当前预览表格的所有数据")
        self.copy_all_btn.setFixedWidth(180)

        # 3. 保存按钮
        self.save_to_excel_btn = QtWidgets.QPushButton("保存到Excel文件")
        self.save_to_excel_btn.setCursor(Qt.PointingHandCursor)
        self.save_to_excel_btn.setToolTip("将预览表格数据保存为Excel文件")
        self.save_to_excel_btn.setFixedWidth(180)

        # 4. 查看日志按钮
        self.show_log_btn = QtWidgets.QPushButton("查看操作日志")
        self.show_log_btn.setCursor(Qt.PointingHandCursor)
        self.show_log_btn.setToolTip("查看完整操作记录")
        self.show_log_btn.setFixedWidth(140)

        # 按钮布局排列
        btn_main_layout.addWidget(self.run_button)
        btn_main_layout.addWidget(self.copy_all_btn)
        btn_main_layout.addWidget(self.save_to_excel_btn)
        btn_main_layout.addWidget(self.show_log_btn)

        main_layout.addLayout(btn_main_layout)

        # ---------------------- 合并结果预览区域（重点保留） ----------------------
        # 预览标题 + 快捷键提示 水平布局
        preview_header_layout = QtWidgets.QHBoxLayout()
        preview_header_layout.setSpacing(10)

        # 预览标题
        self.preview_label = QtWidgets.QLabel("合并结果预览")
        self.preview_label.setStyleSheet("font-weight: bold; font-size: 14px;")

        # 快捷键提示
        self.shortcut_label = QtWidgets.QLabel("快捷键：Ctrl+A 全选 | Ctrl+C 复制选中内容")
        self.shortcut_label.setStyleSheet("color: #666; font-size: 12px;")

        # 标题居左，快捷键居右
        preview_header_layout.addWidget(self.preview_label)
        preview_header_layout.addStretch(1)
        preview_header_layout.addWidget(self.shortcut_label)

        main_layout.addLayout(preview_header_layout)

        # 合并结果预览表格（核心组件，确保完整保留）
        self.table_preview = QtWidgets.QTableWidget()
        self.table_preview.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table_preview.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table_preview.setSelectionBehavior(QAbstractItemView.SelectItems)

        # 表格样式美化
        self.table_preview.setShowGrid(True)
        self.table_preview.setGridStyle(Qt.PenStyle.DotLine)
        self.table_preview.setStyleSheet("""
            QTableWidget {
                border: 1px solid #E0E0E0;
                border-radius: 4px;
                gridline-color: #F0F0F0;
            }
            QHeaderView::section {
                background-color: #F5F7FA;
                color: #333333;
                border: 1px solid #E0E0E0;
                font-size: 11px;
                font-weight: 500;
                padding: 8px 12px;
                text-align: center;
            }
            QTableWidget::item {
                padding: 6px 12px;
                border: 1px solid #F0F0F0;
                font-size: 10.5px;
                color: #444444;
            }
            QTableWidget::item:even {
                background-color: #FFFFFF;
            }
            QTableWidget::item:odd {
                background-color: #FAFBFF;
            }
            QTableWidget::item:selected {
                background-color: #D4E6FC;
                color: #1A56DB;
            }
        """)

        # 表头设置
        self.table_preview.horizontalHeader().setStretchLastSection(True)
        self.table_preview.horizontalHeader().setSectionResizeMode(
            QtWidgets.QHeaderView.ResizeToContents
        )

        # 表格占据主窗口大部分空间（权重1），确保足够大
        main_layout.addWidget(self.table_preview, 1)

        # ---------------------- 状态栏 ----------------------
        self.status_bar = self.statusBar()
        self.status_bar.setStyleSheet("font-size: 12px; padding: 4px 8px;")
        self.status_bar.showMessage("就绪 - 请选择文件并设置参数", 5000)

    def bind_signals(self):
        """绑定信号与槽函数"""
        # 文件选择按钮
        self.first_file_button.clicked.connect(self.browse_first_file)
        self.second_file_button.clicked.connect(self.browse_second_file)

        # 合并按钮
        self.run_button.clicked.connect(self.run_merge)

        # 日志相关
        self.log_update_signal.connect(self.update_log)
        self.show_log_btn.clicked.connect(self.open_log_dialog)

        # 表格更新
        self.update_table_signal.connect(self.update_table_preview)

        # 复制与保存按钮
        self.copy_all_btn.clicked.connect(self.copy_all_table)
        self.save_to_excel_btn.clicked.connect(self.save_to_excel)

        # 快捷键
        self.select_all_shortcut = QShortcut(QKeySequence("Ctrl+A"), self.table_preview)
        self.copy_shortcut = QShortcut(QKeySequence("Ctrl+C"), self.table_preview)
        self.select_all_shortcut.activated.connect(self.table_preview.selectAll)
        self.copy_shortcut.activated.connect(self.copy_selected)

    def open_log_dialog(self):
        """打开日志弹窗"""
        log_dialog = LogDialog(self, self.log_buffer)
        log_dialog.exec_()

    def browse_first_file(self):
        """浏览并选择目标文件"""
        options = QtWidgets.QFileDialog.Options()
        options |= QtWidgets.QFileDialog.ReadOnly
        file_path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "选择目标文件", "", "Excel文件 (*.xlsx);;所有文件 (*)", options=options
        )
        if file_path:
            self.first_file_input.setText(file_path)
            log_msg = f"已选择目标文件：{os.path.basename(file_path)}"
            self.log_update_signal.emit(log_msg)

    def browse_second_file(self):
        """浏览并选择数据文件"""
        options = QtWidgets.QFileDialog.Options()
        options |= QtWidgets.QFileDialog.ReadOnly
        file_path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "选择数据文件", "", "Excel文件 (*.xlsx);;所有文件 (*)", options=options
        )
        if file_path:
            self.second_file_input.setText(file_path)
            log_msg = f"已选择数据文件：{os.path.basename(file_path)}"
            self.log_update_signal.emit(log_msg)

    def run_merge(self):
        """执行数据合并操作"""
        # 获取输入参数
        first_file = self.first_file_input.text().strip()
        second_file = self.second_file_input.text().strip()

        # 校验文件路径
        if not first_file or not os.path.exists(first_file):
            error_msg = "错误：请选择有效的目标文件！"
            self.log_update_signal.emit(error_msg)
            return
        if not second_file or not os.path.exists(second_file):
            error_msg = "错误：请选择有效的数据文件！"
            self.log_update_signal.emit(error_msg)
            return

        # 转换行号（输入行号-1）
        first_row = self.first_row_input.value() - 1
        second_row = self.second_row_input.value() - 1

        # 获取列名参数
        m_id1 = self.first_col_name_input.text().strip()
        m_id2 = self.second_col_name_input.text().strip()
        m_price = self.first_data_col_input.text().strip()
        m_price2 = self.second_data_col_input.text().strip()

        # 校验列名参数
        if not all([m_id1, m_id2, m_price, m_price2]):
            error_msg = "错误：请填写完整的列名参数！"
            self.log_update_signal.emit(error_msg)
            return

        try:
            # 读取目标文件
            log_msg = "正在读取目标文件..."
            self.log_update_signal.emit(log_msg)

            test1_xls = pd.ExcelFile(first_file)
            sheet_names = test1_xls.sheet_names
            log_msg = f"目标文件包含 {len(sheet_names)} 个工作表"
            self.log_update_signal.emit(log_msg)

            # 处理每个工作表
            for idx, sheet_name in enumerate(sheet_names, 1):
                log_msg = f"正在处理工作表 {idx}/{len(sheet_names)}：{sheet_name}"
                self.log_update_signal.emit(log_msg)

                # 读取目标工作表数据
                test1_df = test1_xls.parse(sheet_name, header=first_row)
                if test1_df.empty:
                    log_msg = f"警告：{sheet_name} 为空，已跳过"
                    self.log_update_signal.emit(log_msg)
                    continue

                # 清理列名空格
                test1_df.columns = test1_df.columns.str.strip()

                # 读取数据文件所有工作表并合并
                log_msg = "正在读取数据文件工作表..."
                self.log_update_signal.emit(log_msg)

                test2_xls = pd.ExcelFile(second_file)
                merged_df1 = pd.DataFrame(columns=[m_id1, m_price2])

                for sheet_name2 in test2_xls.sheet_names:
                    test2_df = pd.read_excel(test2_xls, sheet_name2, header=second_row)
                    test2_df.columns = test2_df.columns.str.strip()

                    # 校验数据文件列名
                    if m_id2 not in test2_df.columns or m_price2 not in test2_df.columns:
                        error_msg = f"错误：数据文件工作表 {sheet_name2} 缺少列 {m_id2} 或 {m_price2}"
                        self.log_update_signal.emit(error_msg)
                        return

                    merged_df1 = pd.concat([merged_df1, test2_df[[m_id2, m_price2]]], ignore_index=True)

                # 校验目标文件列名
                if m_id1 not in test1_df.columns or m_price not in test1_df.columns:
                    error_msg = f"错误：目标工作表 {sheet_name} 缺少列 {m_id1} 或 {m_price}"
                    self.log_update_signal.emit(error_msg)
                    return

                # 合并数据
                log_msg = f"正在合并 {sheet_name} 数据..."
                self.log_update_signal.emit(log_msg)

                merged_df1.rename(columns={m_id2: m_id1}, inplace=True)  # 统一匹配列名
                merged_df_final = test1_df.merge(merged_df1, on=m_id1, how='left')

                # 更新价格列
                test1_df[m_price] = merged_df_final[m_price2]

                log_msg = f"{sheet_name} 处理完成"
                self.log_update_signal.emit(log_msg)

                # 更新当前合并结果并刷新表格预览（关键：确保预览表格被更新）
                self.current_merged_df = test1_df.copy()
                self.update_table_signal.emit(self.current_merged_df)

            log_msg = "✅ 所有工作表处理完成！"
            self.log_update_signal.emit(log_msg)

        except KeyError as e:
            error_msg = f"错误（KeyError）：{str(e)}，请检查列名参数"
            self.log_update_signal.emit(error_msg)
        except pd.errors.EmptyDataError:
            error_msg = "错误：读取的文件为空，请检查文件有效性"
            self.log_update_signal.emit(error_msg)
        except Exception as e:
            error_msg = f"未知错误：{str(e)}"
            self.log_update_signal.emit(error_msg)
        finally:
            # 确保文件被关闭
            if 'test1_xls' in locals() and test1_xls:
                test1_xls.close()
            if 'test2_xls' in locals() and test2_xls:
                test2_xls.close()

    def update_table_preview(self, df):
        """更新合并结果预览表格（核心方法）"""
        # 清空表格
        self.table_preview.setRowCount(0)
        self.table_preview.setColumnCount(0)

        # 设置列名
        columns = list(df.columns)
        self.table_preview.setColumnCount(len(columns))
        self.table_preview.setHorizontalHeaderLabels(columns)

        # 设置行数据
        row_count = len(df)
        self.table_preview.setRowCount(row_count)

        # 填充数据
        for row_idx in range(row_count):
            for col_idx in range(len(columns)):
                cell_value = df.iloc[row_idx, col_idx]

                # 处理单元格值
                if pd.isna(cell_value):
                    cell_text = ""
                elif col_idx == 0 and isinstance(cell_value, (int, float)):
                    cell_text = f"{int(cell_value)}"
                elif isinstance(cell_value, (int, float)):
                    cell_text = f"{cell_value:.2f}"
                else:
                    cell_text = str(cell_value)

                # 创建表格项并设置
                item = QTableWidgetItem(cell_text)
                if isinstance(cell_value, (int, float)) and not pd.isna(cell_value):
                    item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                else:
                    item.setTextAlignment(Qt.AlignCenter)
                self.table_preview.setItem(row_idx, col_idx, item)

        # 日志提示
        log_msg = f"表格预览已更新：{row_count} 行 × {len(columns)} 列"
        self.log_update_signal.emit(log_msg)

    def update_log(self, message):
        """更新日志"""
        timestamp = QtCore.QDateTime.currentDateTime().toString("[yyyy-MM-dd HH:mm:ss]")
        log_entry = f"{timestamp} {message}\n"
        self.log_buffer += log_entry
        self.status_bar.showMessage(message, 5000)

    def copy_all_table(self):
        """复制全部表格内容"""
        if self.current_merged_df is None or self.current_merged_df.empty:
            msg = "提示：无数据可复制（表格为空）"
            self.log_update_signal.emit(msg)
            return

        table_text = self.current_merged_df.to_csv(sep='\t', index=False, na_rep='')
        clipboard = QtWidgets.QApplication.clipboard()
        clipboard.setText(table_text)

        msg = f"✅ 已复制全部数据（{len(self.current_merged_df)} 行）"
        self.log_update_signal.emit(msg)

    def copy_selected(self):
        """复制选中的单元格、行或列"""
        selected_items = self.table_preview.selectedItems()
        if not selected_items:
            msg = "提示：请先选中要复制的内容"
            self.log_update_signal.emit(msg)
            return

        data = {}
        min_row = float('inf')
        max_row = -float('inf')
        min_col = float('inf')
        max_col = -float('inf')

        for item in selected_items:
            row = item.row()
            col = item.column()
            if row not in data:
                data[row] = {}
            data[row][col] = item.text()

            min_row = min(min_row, row)
            max_row = max(max_row, row)
            min_col = min(min_col, col)
            max_col = max(max_col, col)

        headers = []
        for col in range(min_col, max_col + 1):
            header_item = self.table_preview.horizontalHeaderItem(col)
            headers.append(header_item.text() if header_item else f"列{col + 1}")

        copy_text = '\t'.join(headers) + '\n'
        for row in range(min_row, max_row + 1):
            row_data = []
            for col in range(min_col, max_col + 1):
                row_data.append(data.get(row, {}).get(col, ""))
            copy_text += '\t'.join(row_data) + '\n'

        clipboard = QtWidgets.QApplication.clipboard()
        clipboard.setText(copy_text.strip())

        selected_rows = max_row - min_row + 1
        selected_cols = max_col - min_col + 1
        msg = f"✅ 已复制选中区域（{selected_rows} 行 × {selected_cols} 列）"
        self.log_update_signal.emit(msg)

    def save_to_excel(self):
        """将当前预览的表格数据保存为Excel文件"""
        if self.current_merged_df is None or self.current_merged_df.empty:
            msg = "提示：无数据可保存（表格为空）"
            self.log_update_signal.emit(msg)
            return

        default_dir = os.path.join(os.path.expanduser("~"), "Desktop")
        if not os.path.exists(default_dir):
            default_dir = os.path.expanduser("~")

        file_path, _ = QFileDialog.getSaveFileName(
            self, "保存Excel文件", default_dir, "Excel文件 (*.xlsx);;所有文件 (*)"
        )

        if file_path:
            try:
                if not file_path.endswith('.xlsx'):
                    file_path += '.xlsx'
                self.current_merged_df.to_excel(file_path, index=False)
                msg = f"✅ 数据已成功保存到：{file_path}"
                self.log_update_signal.emit(msg)
            except Exception as e:
                error_msg = f"保存失败：{str(e)}"
                self.log_update_signal.emit(error_msg)


def main():
    """主函数"""
    app = QtWidgets.QApplication(sys.argv)
    app.setStyle("Fusion")
    window = ExcelMergerApp()
    window.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
