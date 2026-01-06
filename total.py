import sys
from PySide6.QtWidgets import QApplication, QMainWindow, QPushButton, QFileDialog, QVBoxLayout, QWidget, QLabel, QListWidget, QHBoxLayout, QMessageBox, QTableWidget, QTableWidgetItem, QDialog, QTextEdit
from PySide6.QtGui import QKeySequence, QShortcut
from PySide6.QtCore import Qt, Signal, QDateTime
import pandas as pd
import os
import calendar
from datetime import datetime

class LogDialog(QDialog):
    def __init__(self, parent=None, log_text=""):
        super().__init__(parent)
        self.setWindowTitle("操作日志")
        self.resize(800, 500)
        layout = QVBoxLayout(self)
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setText(log_text)
        layout.addWidget(self.log_text)
        btn_layout = QHBoxLayout()
        self.clear_btn = QPushButton("清空日志")
        self.copy_btn = QPushButton("复制全部日志")
        btn_layout.addStretch(1)
        btn_layout.addWidget(self.clear_btn)
        btn_layout.addWidget(self.copy_btn)
        layout.addLayout(btn_layout)
        self.clear_btn.clicked.connect(self.clear_log)
        self.copy_btn.clicked.connect(self.copy_log)
    def clear_log(self):
        self.log_text.clear()
        if self.parent():
            self.parent().log_buffer = ""
    def copy_log(self):
        text = self.log_text.toPlainText()
        if text:
            clipboard = QApplication.clipboard()
            clipboard.setText(text)
            if self.parent():
                self.parent().status_bar.showMessage("日志已复制到剪贴板", 3000)


class ExcelProcessor(QMainWindow):
    log_update_signal = Signal(str)
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Excel数据提取工具")
        self.setGeometry(100, 100, 600, 400)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        self.layout = QVBoxLayout()

        # 添加文件操作部分
        self.file_layout = QHBoxLayout()

        self.add_button = QPushButton("添加文件")
        self.add_button.clicked.connect(self.add_file)
        self.file_layout.addWidget(self.add_button)

        self.remove_button = QPushButton("移除文件")
        self.remove_button.clicked.connect(self.remove_file)
        self.file_layout.addWidget(self.remove_button)

        self.layout.addLayout(self.file_layout)

        # 自动识别年份与月份，无需手动输入

        # 显示选择的文件
        self.file_list = QListWidget()
        self.layout.addWidget(self.file_list)

        # 添加提取按钮
        self.extract_button = QPushButton("提取数据")
        self.extract_button.clicked.connect(self.process_selected_files)

        self.copy_button = QPushButton("复制表格内容")
        self.copy_button.clicked.connect(self.copy_all_table)
        self.copy_button.setEnabled(False)

        self.preview_label = QLabel("结果预览（可选中单元格，使用Ctrl+C复制）：")

        self.show_log_btn = QPushButton("查看操作日志")

        self.table_widget = QTableWidget()
        self.table_widget.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table_widget.setSelectionMode(QTableWidget.ExtendedSelection)
        self.table_widget.setSelectionBehavior(QTableWidget.SelectItems)
        self.table_widget.setColumnCount(3)
        self.table_widget.setHorizontalHeaderLabels(["日期", "文件名", "合计数据"])
        self.table_widget.horizontalHeader().setStretchLastSection(True)
        self.layout.addWidget(self.table_widget)

        self.add_copy_shortcut()

        # 添加保存按钮
        self.save_button = QPushButton("保存结果")
        self.save_button.clicked.connect(self.save_results)

        btn_main_layout = QHBoxLayout()
        btn_main_layout.addWidget(self.extract_button)
        btn_main_layout.addWidget(self.copy_button)
        btn_main_layout.addWidget(self.show_log_btn)
        btn_main_layout.addWidget(self.save_button)
        self.layout.addLayout(btn_main_layout)

        self.layout.addWidget(self.preview_label)

        # 显示结果保存路径
        self.save_path_label = QLabel("保存路径：")
        self.layout.addWidget(self.save_path_label)

        self.central_widget.setLayout(self.layout)

        # 存储选择的文件路径
        self.selected_files = []
        self.results = []
        self.log_buffer = ""
        self.status_bar = self.statusBar()
        self.status_bar.showMessage("就绪")
        self.log_update_signal.connect(self.update_log)
        self.show_log_btn.clicked.connect(self.open_log_dialog)

        # # 窗口占屏幕的比例
        # self.scale = 0.6
        # # 调整窗口大小并居中
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
        shortcut_select_all = QShortcut(QKeySequence("Ctrl+A"), self.table_widget)
        shortcut_select_all.activated.connect(self.table_widget.selectAll)
        shortcut_copy = QShortcut(QKeySequence("Ctrl+C"), self.table_widget)
        shortcut_copy.activated.connect(self.copy_selected_cells)
    def add_file(self):
        file_dialog = QFileDialog()
        file_dialog.setNameFilter("Excel files (*.xlsx)")
        file_dialog.setFileMode(QFileDialog.ExistingFiles)
        if file_dialog.exec():
            selected_files = file_dialog.selectedFiles()
            for file_path in selected_files:
                if file_path not in self.selected_files:
                    self.selected_files.append(file_path)
                    self.file_list.addItem(os.path.basename(file_path))
                else:
                    QMessageBox.warning(self, "Warning", f"{os.path.basename(file_path)} 已经添加过了！")

    def remove_file(self):
        selected_items = self.file_list.selectedItems()
        for item in selected_items:
            row = self.file_list.row(item)
            self.selected_files.pop(row)
            self.file_list.takeItem(row)

    def process_excel(self, file_path):
        # 读取Excel文件
        df = pd.read_excel(file_path, header=None)

        # 找到合计所在的行
        total_row_index = None
        for i, row in df.iterrows():
            if '合计' in str(row.values):
                total_row_index = i
                break

        if total_row_index is not None and total_row_index < len(df) - 1:
            # 合计的下一行就是需要的数据
            data_row_index = total_row_index + 1
            data = df.iloc[data_row_index]

            # 获取合计所在列的数据
            total_column_index = row.index[row.apply(lambda x: '合计' in str(x))][0]
            total_data = data.iloc[total_column_index]

            log_msg = f"在文件 {os.path.basename(file_path)} 中找到合计数据: {total_data}"
            self.log_update_signal.emit(log_msg)

            return os.path.basename(file_path), total_data

    def extract_day_from_filename(self, file_path):
        name = os.path.basename(file_path)
        base = os.path.splitext(name)[0]
        import re
        m = re.search(r"(?P<y>\d{4})[\-_/\.](?P<m>\d{1,2})[\-_/\.](?P<d>\d{1,2})", base)
        if m:
            return int(m.group("d"))
        m = re.search(r"(?:(\d{4})年)?(?:(\d{1,2})月)?(\d{1,2})[日号]", base)
        if m:
            return int(m.group(3))
        m = re.search(r"(?<!\d)(\d{1,2})(?:[日号])(?!\d)", base)
        if m:
            return int(m.group(1))
        m = re.search(r"[\-_.](\d{1,2})(?!.*\d)", base)
        if m:
            return int(m.group(1))
        m = re.search(r"(\d{1,2})(?=\.[^.]+$)", name)
        if m:
            return int(m.group(1))
        return None

    def extract_year_month_from_filename(self, file_path):
        name = os.path.basename(file_path)
        base = os.path.splitext(name)[0]
        import re
        m = re.search(r"(?P<y>\d{4})[\-_/\.](?P<m>\d{1,2})[\-_/\.]\d{1,2}", base)
        if m:
            return int(m.group("y")), int(m.group("m"))
        m = re.search(r"(?P<y>\d{4})[\-_/\.](?P<m>\d{1,2})(?!.*\d)", base)
        if m:
            return int(m.group("y")), int(m.group("m"))
        m = re.search(r"(?P<y>\d{4})年(?P<m>\d{1,2})月", base)
        if m:
            return int(m.group("y")), int(m.group("m"))
        m = re.search(r"(?P<m>\d{1,2})月\d{1,2}[日号]", base)
        if m:
            return None, int(m.group("m"))
        return None

    def extract_year_month_from_path(self, file_path):
        import re
        dir_name = os.path.basename(os.path.dirname(file_path))
        m = re.search(r"(?P<y>\d{4})[\-_/\.](?P<m>\d{1,2})", dir_name)
        if m:
            return int(m.group("y")), int(m.group("m"))
        m = re.search(r"(?P<y>\d{4})年(?P<m>\d{1,2})月", dir_name)
        if m:
            return int(m.group("y")), int(m.group("m"))
        return None

    def determine_year_month(self, files):
        from collections import Counter
        ym_counter = Counter()
        # 尝试从文件名提取
        for fp in files:
            ym = self.extract_year_month_from_filename(fp)
            if isinstance(ym, tuple) and ym[0] is not None and ym[1] is not None:
                ym_counter[ym] += 1
        if ym_counter:
            return max(ym_counter.items(), key=lambda x: x[1])[0]
        # 尝试从上级文件夹提取
        for fp in files:
            ym = self.extract_year_month_from_path(fp)
            if ym:
                ym_counter[ym] += 1
        if ym_counter:
            return max(ym_counter.items(), key=lambda x: x[1])[0]
        # 回退到按文件修改时间的年月（取出现次数最多的）
        ym_counter = Counter()
        for fp in files:
            try:
                dt = datetime.fromtimestamp(os.path.getmtime(fp))
                ym_counter[(dt.year, dt.month)] += 1
            except Exception:
                continue
        if ym_counter:
            return max(ym_counter.items(), key=lambda x: x[1])[0]
        # 最后兜底为当前年月
        today = QDateTime.currentDateTime().date()
        return today.year(), today.month()

    def format_date(self, y, m, d):
        return f"{y:04d}-{m:02d}-{d:02d}"

    def process_selected_files(self):
        self.results = []
        y, m = self.determine_year_month(self.selected_files)
        day_entries = {}
        for file_path in self.selected_files:
            day = self.extract_day_from_filename(file_path)
            if day is None:
                self.log_update_signal.emit(f"无法从文件名识别日期：{os.path.basename(file_path)}")
                continue
            result = self.process_excel(file_path)
            if result:
                if day not in day_entries:
                    day_entries[day] = (os.path.basename(file_path), result[1])
        last_day = calendar.monthrange(y, m)[1]
        for d in range(1, last_day + 1):
            date_str = self.format_date(y, m, d)
            if d in day_entries:
                fname, total_data = day_entries[d]
                self.results.append((date_str, fname, total_data))
            else:
                self.results.append((date_str, "", ""))
        self.update_table_preview()
        self.copy_button.setEnabled(bool(self.results))
        self.log_update_signal.emit(f"共提取 {len(self.results)} 条记录")

    def save_results(self):
        if not self.results:
            QMessageBox.warning(self, "Warning", "没有可保存的结果！")
            return

        save_path, _ = QFileDialog.getSaveFileName(self, "保存结果", "", "Excel files (*.xlsx)")
        if save_path:
            df = pd.DataFrame(self.results, columns=['日期', '文件名', '合计数据'])
            df.to_excel(save_path, index=False)
            self.save_path_label.setText(f"保存路径：{save_path}")

    def update_table_preview(self):
        self.table_widget.clearContents()
        self.table_widget.setRowCount(0)
        self.table_widget.setColumnCount(3)
        self.table_widget.setHorizontalHeaderLabels(["日期", "文件名", "合计数据"])
        if not self.results:
            return
        self.table_widget.setRowCount(len(self.results))
        for row_idx, (date_str, fname, total_data) in enumerate(self.results):
            item_date = QTableWidgetItem(str(date_str))
            item_date.setTextAlignment(Qt.AlignCenter)
            self.table_widget.setItem(row_idx, 0, item_date)
            item_name = QTableWidgetItem(str(fname))
            item_name.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            self.table_widget.setItem(row_idx, 1, item_name)
            item_total = QTableWidgetItem(str(total_data))
            if isinstance(total_data, (int, float)):
                item_total.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            else:
                item_total.setTextAlignment(Qt.AlignCenter)
            self.table_widget.setItem(row_idx, 2, item_total)
        self.table_widget.resizeColumnsToContents()
        self.log_update_signal.emit(f"表格预览已更新：{len(self.results)} 行 × 3 列")

    def copy_selected_cells(self):
        selected_items = self.table_widget.selectedItems()
        if not selected_items:
            return
        selected_items.sort(key=lambda x: (x.row(), x.column()))
        rows = sorted(set(item.row() for item in selected_items))
        cols = sorted(set(item.column() for item in selected_items))
        clipboard_text = ""
        for row in rows:
            row_data = []
            for col in cols:
                for item in selected_items:
                    if item.row() == row and item.column() == col:
                        row_data.append(item.text())
                        break
            clipboard_text += "\t".join(row_data) + "\n"
        clipboard_text = clipboard_text.rstrip("\n")
        clipboard = QApplication.clipboard()
        clipboard.setText(clipboard_text)

    def copy_all_table(self):
        if not self.results:
            return
        header = "日期\t文件名\t合计数据\n"
        lines = []
        for date_str, fname, total_data in self.results:
            lines.append(f"{date_str}\t{fname}\t{total_data}")
        clipboard = QApplication.clipboard()
        clipboard.setText(header + "\n".join(lines))
        self.log_update_signal.emit(f"已复制全部数据（{len(self.results)} 行）")

    def update_log(self, message):
        timestamp = QDateTime.currentDateTime().toString("[yyyy-MM-dd HH:mm:ss]")
        entry = f"{timestamp} {message}\n"
        self.log_buffer += entry
        self.status_bar.showMessage(message, 5000)

    def open_log_dialog(self):
        dlg = LogDialog(self, self.log_buffer)
        dlg.exec_()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelProcessor()
    window.show()
    sys.exit(app.exec())
