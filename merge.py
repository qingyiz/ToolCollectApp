# import pandas as pd
#
# # 读取test1.xlsx和test2.xlsx
# test1_file = '五矿10-19生鲜采购.xlsx'
# test2_file = '后浪餐厅生鲜食材定价(10..16-10.31）-新版.xlsx'
#
# test1_df = pd.read_excel(test1_file, header=2)
# test2_df = pd.read_excel(test2_file, header=0)
#
# # 去除列名中的额外空格
# test1_df.columns = test1_df.columns.str.strip()
# test2_df.columns = test2_df.columns.str.strip()
#
#
# # 合并数据
# merged_df = test1_df.merge(test2_df[['商品名称', 'TG']], on='商品名称', how='left')
#
# # 更新test1中的'price'列
# test1_df['价格'] = merged_df['TG']
#
# # 保存结果到一个新的Excel文件
# output_file = 'merged_test2.xlsx'
# test1_df.to_excel(output_file, index=False)
#
# print("数据已合并并保存到新的Excel文件:", output_file)
import datetime
import os
# 以下是可以读取第二个execl中有多个工作表的情况
# import pandas as pd
#
# m_id = '商品名称'
# m_price = '价格'
# m_price2 = 'TG'
# header1 = 2
# header2 = 0
#
#
# # 读取test1.xlsx的数据
# test1_file = '五矿10-19生鲜采购.xlsx'
# test1_df = pd.read_excel(test1_file, header=2)
#
# # 指定列名所在的行号
# test1_df.columns = test1_df.columns.str.strip()
#
# # 创建一个空的DataFrame来存储合并后的数据
# merged_df = pd.DataFrame(columns=[m_id, m_price])
#
# # 读取test2.xlsx中的所有工作表
# test2_file = '后浪餐厅生鲜食材定价(10..16-10.31）-新版.xlsx'
# xls = pd.ExcelFile(test2_file)
# for sheet_name in xls.sheet_names:
#     # 读取每个工作表
#     test2_df = pd.read_excel(xls, sheet_name, header=0)
#     test2_df.columns = test2_df.columns.str.strip()
#
#     # 合并数据到merged_df
#     merged_df = pd.concat([merged_df, test2_df[[m_id, m_price2]]])
#
# # 合并数据
# merged_df = test1_df.merge(merged_df, on=m_id, how='left')
#
# # 更新test1中的'price'列并保留两位小数
# test1_df[m_price] = merged_df[m_price2]
#
# # 保存结果到一个新的Excel文件
# output_file = 'merged_test1.xlsx'
# test1_df.to_excel(output_file, index=False)
#
# print("数据已合并并保存到新的Excel文件:", output_file)



#
# import pandas as pd
#
# m_id1 = '商品名称'
# m_id2 = '商品名称'
# m_price = '价格'
# m_price2 = 'TG'
# header1 = 2
# header2 = 0
#
# output_file = 'output.xlsx'
# output_xls = pd.ExcelWriter(output_file)
#
#
# # 读取test1.xlsx的数据
# test1_file = '10月22日到货生鲜采购单.xlsx'
# # test1_file = "10月19日到货生鲜采购单.xlsx"
# test1_xls = pd.ExcelFile(test1_file)
#
# for sheet_name in test1_xls.sheet_names:
#     print(sheet_name)
#     test1_df = test1_xls.parse(sheet_name, header=header1)
#     test1_df.columns = test1_df.columns.str.strip()
#     test2_file = '后浪餐厅生鲜食材定价(10..16-10.31）-新版.xlsx'
#     test2_xls = pd.ExcelFile(test2_file)
#
#     # 创建一个空的DataFrame来存储合并后的数据
#     merged_df1 = pd.DataFrame(columns=[m_id1, m_price])
#
#     # 读取test2.xlsx中的所有工作表
#     test2_file = '后浪餐厅生鲜食材定价(10..16-10.31）-新版.xlsx'
#     xls = pd.ExcelFile(test2_file)
#     for sheet_name2 in xls.sheet_names:
#         # 读取每个工作表
#         test2_df = pd.read_excel(xls, sheet_name2, header=header2)
#         test2_df.columns = test2_df.columns.str.strip()
#
#         # 合并数据到merged_df
#         merged_df1 = pd.concat([merged_df1, test2_df[[m_id1, m_price2]]])
#
#     # 合并数据
#     merged_df1 = test1_df.merge(merged_df1, on=m_id1, how='left')
#
#     # 更新test1中的'price'列并保留两位小数
#     test1_df[m_price] = merged_df1[m_price2]
#
#     # 保存结果到一个新的Excel文件
#
#     print("save")
#     test1_df.to_excel(output_xls, sheet_name=sheet_name, index=False)
#
# print("数据已合并并保存到新的Excel文件:", output_file)
# output_xls.close()


import sys
import pandas as pd
from PySide6 import QtWidgets, QtGui, QtCore
from PySide6.QtGui import QIcon


class ExcelMergerApp(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("数据合并")
        # 创建标签和输入字段
        target_file_label = QtWidgets.QLabel("目标文件:")
        self.first_file_input = QtWidgets.QLineEdit()
        self.first_file_button = QtWidgets.QPushButton("浏览")

        self.second_file_label = QtWidgets.QLabel("数据文件:")
        self.second_file_input = QtWidgets.QLineEdit()
        self.second_file_button = QtWidgets.QPushButton("浏览")

        self.parameters_label = QtWidgets.QLabel("参数设置")
        self.first_row_label = QtWidgets.QLabel("目标文件列名所在行号:")
        self.first_row_input = QtWidgets.QSpinBox()
        self.first_row_input.setValue(3)
        self.first_col_name_label = QtWidgets.QLabel("需要匹配的列名:")
        self.first_col_name_input = QtWidgets.QLineEdit()
        self.first_col_name_input.setText("商品名称")
        self.first_data_col_label = QtWidgets.QLabel("目标文件需要填入数据的列名:")
        self.first_data_col_input = QtWidgets.QLineEdit()
        self.first_data_col_input.setText("价格")

        self.second_row_label = QtWidgets.QLabel("数据文件列名所在行号:")
        self.second_row_input = QtWidgets.QSpinBox()
        self.second_row_input.setValue(1)
        self.second_col_name_label = QtWidgets.QLabel("数据文件需要匹配的列名:")
        self.second_col_name_input = QtWidgets.QLineEdit()
        self.second_col_name_input.setText("商品名称")
        self.second_data_col_label = QtWidgets.QLabel("数据文件数据来源列名:")
        self.second_data_col_input = QtWidgets.QLineEdit()
        self.second_data_col_input.setText("TG")

        self.run_button = QtWidgets.QPushButton("执行合并")

        self.log_label = QtWidgets.QLabel("日志:")
        self.log_text = QtWidgets.QTextEdit()
        self.log_text.setReadOnly(True)

        # 创建布局
        layout = QtWidgets.QGridLayout()
        layout.addWidget(target_file_label, 0, 0)
        layout.addWidget(self.first_file_input, 0, 1)
        layout.addWidget(self.first_file_button, 0, 2)

        layout.addWidget(self.second_file_label, 0, 3)
        layout.addWidget(self.second_file_input, 0, 4)
        layout.addWidget(self.second_file_button, 0, 5)

        layout.addWidget(self.parameters_label, 1, 0, 1, 6)

        layout.addWidget(self.first_row_label, 2, 0)
        layout.addWidget(self.first_row_input, 2, 1)
        layout.addWidget(self.first_col_name_label, 3, 0)
        layout.addWidget(self.first_col_name_input, 3, 1)
        layout.addWidget(self.first_data_col_label, 4, 0)
        layout.addWidget(self.first_data_col_input, 4, 1)

        layout.addWidget(self.second_row_label, 2, 3)
        layout.addWidget(self.second_row_input, 2, 4)
        layout.addWidget(self.second_col_name_label, 3, 3)
        layout.addWidget(self.second_col_name_input, 3, 4)
        layout.addWidget(self.second_data_col_label, 4, 3)
        layout.addWidget(self.second_data_col_input, 4, 4)

        layout.addWidget(self.run_button, 5, 0, 1, 6)
        layout.addWidget(self.log_label, 6, 0)
        layout.addWidget(self.log_text, 7, 0, 1, 6)

        self.setLayout(layout)

        # 连接按钮到函数
        self.first_file_button.clicked.connect(self.browse_first_file)
        self.second_file_button.clicked.connect(self.browse_second_file)
        self.run_button.clicked.connect(self.run_merge)

    def browse_first_file(self):
        options = QtWidgets.QFileDialog.Options()
        options |= QtWidgets.QFileDialog.ReadOnly
        file_dialog = QtWidgets.QFileDialog()
        first_file, _ = file_dialog.getOpenFileName(self, "选择目标文件", "", "Excel文件 (*.xlsx);;所有文件 (*)", options=options)
        if first_file:
            self.first_file_input.setText(first_file)

    def browse_second_file(self):
        options = QtWidgets.QFileDialog.Options()
        options |= QtWidgets.QFileDialog.ReadOnly
        file_dialog = QtWidgets.QFileDialog()
        second_file, _ = file_dialog.getOpenFileName(self, "选择第数据文件", "", "Excel文件 (*.xlsx);;所有文件 (*)", options=options)
        if second_file:
            self.second_file_input.setText(second_file)

    def run_merge(self):
        # 获取输入值
        first_file = self.first_file_input.text()
        first_row = self.first_row_input.value() - 1
        first_col_name = self.first_col_name_input.text()
        first_data_col = self.first_data_col_input.text()
        second_file = self.second_file_input.text()
        second_row = self.second_row_input.value() - 1
        second_col_name = self.second_col_name_input.text()
        second_data_col = self.second_data_col_input.text()

        # 在此处插入您的现有代码（使用参数化值）
        m_id1 = first_col_name
        m_id2 = second_col_name
        m_price = first_data_col
        m_price2 = second_data_col
        header1 = first_row
        header2 = second_row
        # 获取用户主目录
        home_dir = os.path.expanduser("~")
        # 拼接桌面路径
        desktop_path = os.path.join(home_dir, "Desktop")
        output_file = os.path.join(desktop_path, "output.xlsx")
        output_xls = pd.ExcelWriter(output_file)
        test1_xls = None
        try:
            test1_xls = pd.ExcelFile(first_file)
        except KeyError as e:
            self.log(f"发生 KeyError 异常: {e}")

        for sheet_name in test1_xls.sheet_names:
            test1_df = test1_xls.parse(sheet_name, header=header1)
            if test1_df.empty:
                self.log(f"{sheet_name}为空,已跳过")
                continue
            test1_df.columns = test1_df.columns.str.strip()

            test2_xls = pd.ExcelFile(second_file)
            merged_df1 = pd.DataFrame(columns=[m_id1, m_price])
            for sheet_name2 in test2_xls.sheet_names:
                test2_df = pd.read_excel(test2_xls, sheet_name2, header=header2)
                test2_df.columns = test2_df.columns.str.strip()
                try:
                    merged_df1 = pd.concat([merged_df1, test2_df[[m_id1, m_price2]]])
                except KeyError as e:
                    # 处理异常的代码，例如输出错误信息或者提供默认值
                    self.log(f"发生 KeyError 异常: {e}")
                    self.log("请检查【数据文件】【行号】以及【列名】是否填写正确")
                    # 可以提供默认值或者创建一个包含所需列的DataFrame
                    return
            try:
                merged_df1 = test1_df.merge(merged_df1, on=m_id1, how='left')
            except pd.errors.MergeError as e:
                # 处理 MergeError 异常
                print(f"发生 MergeError 异常: {e}")
            except KeyError as e:
                # 处理 KeyError 异常
                self.log(f"发生 KeyError 异常: {e}")
                self.log(f"请检查【目标文件】工作表名：【{sheet_name}】【行号】以及【列名】是否填写正确")
                return
            except ValueError as e:
                # 处理 ValueError 异常
                print(f"发生 ValueError 异常: {e}")
                return
            test1_df[m_price] = merged_df1[m_price2]
            test1_df.to_excel(output_xls, sheet_name=sheet_name, index=False)
            self.log(f"{sheet_name } 合并成功")

        output_xls.close()
        self.log("数据已合并并保存到新的Excel文件:" + output_file)

    def log(self, message):
        current_time = datetime.datetime.now()
        timestamp = current_time.strftime("[%Y-%m-%d %H:%M:%S] ")
        log_message = timestamp + message
        current_text = self.log_text.toPlainText()
        self.log_text.setPlainText(current_text + log_message + "\n")
        # 自动滚动到底部
        self.log_text.verticalScrollBar().setValue(self.log_text.verticalScrollBar().maximum())

def main():
    app = QtWidgets.QApplication([])
    window = ExcelMergerApp()
    window.setWindowTitle("Excel数据合并")
    # window.setWindowIcon(QIcon("icon.svg"))
    window.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()