import sys
from PySide6.QtWidgets import QApplication, QMainWindow, QPushButton, QFileDialog, QVBoxLayout, QWidget, QLabel, QListWidget, QHBoxLayout, QMessageBox, QTextEdit
import pandas as pd
import os

class ExcelProcessor(QMainWindow):
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

        # 显示选择的文件
        self.file_list = QListWidget()
        self.layout.addWidget(self.file_list)

        # 添加提取按钮
        self.extract_button = QPushButton("提取数据")
        self.extract_button.clicked.connect(self.process_selected_files)
        self.layout.addWidget(self.extract_button)

        # 显示结果
        self.result_textbox = QTextEdit()
        self.layout.addWidget(self.result_textbox)

        # 添加保存按钮
        self.save_button = QPushButton("保存结果")
        self.save_button.clicked.connect(self.save_results)
        self.layout.addWidget(self.save_button)

        # 显示结果保存路径
        self.save_path_label = QLabel("保存路径：")
        self.layout.addWidget(self.save_path_label)

        self.central_widget.setLayout(self.layout)

        # 存储选择的文件路径
        self.selected_files = []
        self.results = []

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

            # 输出结果到界面
            result_text = f"在文件 {os.path.basename(file_path)} 中找到合计数据: {total_data}\n"
            self.result_textbox.append(result_text)

            return os.path.basename(file_path), total_data

    def process_selected_files(self):
        self.result_textbox.clear()
        self.results = []
        for file_path in self.selected_files:
            result = self.process_excel(file_path)
            if result:
                self.results.append(result)

    def save_results(self):
        if not self.results:
            QMessageBox.warning(self, "Warning", "没有可保存的结果！")
            return

        save_path, _ = QFileDialog.getSaveFileName(self, "保存结果", "", "Excel files (*.xlsx)")
        if save_path:
            df = pd.DataFrame(self.results, columns=['文件名', '合计数据'])
            df.to_excel(save_path, index=False)
            self.save_path_label.setText(f"保存路径：{save_path}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelProcessor()
    window.show()
    sys.exit(app.exec())
