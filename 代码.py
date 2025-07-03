import sys
import numpy as np
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QFileDialog, QVBoxLayout, QWidget, QPushButton, QLabel,
    QTabWidget, QTableWidget, QTableWidgetItem, QHBoxLayout, QMessageBox, QTextEdit, QSplitter
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QTextCursor, QColor, QTextCharFormat
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure

FIELDS = [
    "经度(°)", "纬度(°)", "垂向深度(m)", "井口海拔(m)",
    "X轴分量(nT)", "Y轴分量(nT)", "Z轴分量(nT)"
]
VECTOR_FIELDS = ["X轴分量(nT)", "Y轴分量(nT)", "Z轴分量(nT)"]

def detect_outliers_zscore(data, threshold=3):
    z = np.abs((data - np.mean(data, axis=0)) / np.std(data, axis=0))
    outlier_mask = (z > threshold).any(axis=1)
    return outlier_mask

def weighted_average(data, weights=None):
    if weights is None:
        weights = np.ones(data.shape[0])
    weights = weights / np.sum(weights)
    return np.average(data, axis=0, weights=weights)

def correct_main_data(main_vec, aux_mean, factor=1.0):
    return main_vec + factor * (aux_mean - main_vec)

def table_to_dataframe(table):
    rows = table.rowCount()
    cols = table.columnCount()
    data = []
    for r in range(rows):
        row = []
        for c in range(cols):
            item = table.item(r, c)
            value = item.text() if item else ""
            row.append(value)
        data.append(row)
    df = pd.DataFrame(data, columns=[table.horizontalHeaderItem(i).text() for i in range(cols)])
    for col in FIELDS:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
    return df

def dataframe_to_table(df, table):
    table.setRowCount(df.shape[0])
    table.setColumnCount(df.shape[1])
    table.setHorizontalHeaderLabels(df.columns)
    for r in range(df.shape[0]):
        for c in range(df.shape[1]):
            val = str(df.iat[r, c])
            table.setItem(r, c, QTableWidgetItem(val))

class ColorConsoleStream:
    """A stream object to redirect stdout/stderr to QTextEdit with color support."""
    def __init__(self, text_edit):
        self.text_edit = text_edit

    def write(self, text):
        # Detect type of message and set color
        if '[ERROR]' in text or text.strip().startswith('Traceback'):
            self._write_colored(text, color=QColor('red'))
        elif '[WARN]' in text:
            self._write_colored(text, color=QColor('gold'), bgcolor=QColor(0, 50, 0))  # yellow on dark green
        else:
            self._write_colored(text, color=QColor('black'))
    
    def _write_colored(self, text, color=QColor('black'), bgcolor=None):
        cursor = self.text_edit.textCursor()
        cursor.movePosition(QTextCursor.End)
        char_format = QTextCharFormat()
        char_format.setForeground(color)
        if bgcolor is not None:
            char_format.setBackground(bgcolor)
        cursor.setCharFormat(char_format)
        cursor.insertText(text)
        self.text_edit.setTextCursor(cursor)
        self.text_edit.ensureCursorVisible()
        # Reset to default for next print
        char_format.setForeground(QColor('black'))
        char_format.setBackground(QColor('white'))
        cursor.setCharFormat(char_format)

    def flush(self):
        pass

class MagneticCorrectionGeoApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("地磁测量多点修正可视化（地理坐标版，含内嵌彩色控制台）")
        self.resize(1200, 800)

        # 主Splitter（左数据/右图表及控制台）
        self.splitter_main = QSplitter(Qt.Horizontal, self)
        self.setCentralWidget(self.splitter_main)

        # 左侧：数据&操作
        self.left_widget = QWidget()
        self.left_layout = QVBoxLayout(self.left_widget)

        btn_layout = QHBoxLayout()
        self.btn_load_excel = QPushButton("从Excel读取数据")
        self.btn_load_excel.clicked.connect(self.load_excel)
        btn_layout.addWidget(self.btn_load_excel)

        self.btn_process = QPushButton("开始修正并可视化")
        self.btn_process.clicked.connect(self.process_and_plot)
        btn_layout.addWidget(self.btn_process)

        self.btn_add_main_row = QPushButton("主数据添加一行")
        self.btn_add_main_row.clicked.connect(self.add_main_row)
        btn_layout.addWidget(self.btn_add_main_row)

        self.btn_add_aux_row = QPushButton("辅助数据添加一行")
        self.btn_add_aux_row.clicked.connect(self.add_aux_row)
        btn_layout.addWidget(self.btn_add_aux_row)

        self.left_layout.addLayout(btn_layout)

        self.tabs = QTabWidget()
        self.tab_main = QWidget()
        self.tab_aux = QWidget()
        self.tabs.addTab(self.tab_main, "主数据")
        self.tabs.addTab(self.tab_aux, "辅助数据")
        self.left_layout.addWidget(self.tabs)

        self.table_main = QTableWidget(1, len(FIELDS))
        self.table_main.setHorizontalHeaderLabels(FIELDS)
        tab_main_layout = QVBoxLayout()
        tab_main_layout.addWidget(self.table_main)
        self.tab_main.setLayout(tab_main_layout)

        self.table_aux = QTableWidget(1, len(FIELDS))
        self.table_aux.setHorizontalHeaderLabels(FIELDS)
        tab_aux_layout = QVBoxLayout()
        tab_aux_layout.addWidget(self.table_aux)
        self.tab_aux.setLayout(tab_aux_layout)

        self.label = QLabel(
            "数据格式：经度(°)、纬度(°)、垂向深度(m)、井口海拔(m)、X轴分量(nT)、Y轴分量(nT)、Z轴分量(nT)\n"
            "可手动输入数据，或从Excel读取。"
        )
        self.left_layout.addWidget(self.label)

        self.splitter_main.addWidget(self.left_widget)

        # 右侧：图表和控制台
        self.right_widget = QWidget()
        self.right_layout = QVBoxLayout(self.right_widget)

        self.figure = Figure()
        self.canvas = FigureCanvas(self.figure)
        self.right_layout.addWidget(self.canvas)

        self.console_toggle_btn = QPushButton("显示控制台")
        self.console_toggle_btn.setCheckable(True)
        self.console_toggle_btn.setChecked(False)
        self.console_toggle_btn.clicked.connect(self.toggle_console)
        self.right_layout.addWidget(self.console_toggle_btn)

        self.console_text = QTextEdit()
        self.console_text.setReadOnly(True)
        self.console_text.hide()
        self.right_layout.addWidget(self.console_text)

        self.splitter_main.addWidget(self.right_widget)
        self.splitter_main.setSizes([500, 700])

        # 控制台流重定向
        self.console_stream = ColorConsoleStream(self.console_text)
        self.console_shown = False

        # 隐藏外置控制台
        sys.stdout = open(os.devnull, 'w')
        sys.stderr = open(os.devnull, 'w')

    def add_main_row(self):
        self.table_main.insertRow(self.table_main.rowCount())

    def add_aux_row(self):
        self.table_aux.insertRow(self.table_aux.rowCount())

    def load_excel(self):
        import traceback
        file_path, _ = QFileDialog.getOpenFileName(self, "选择Excel文件", "", "Excel Files (*.xlsx *.xls)")
        if not file_path:
            return
        try:
            xls = pd.ExcelFile(file_path)
            main_df = pd.read_excel(xls, xls.sheet_names[0])
            aux_df = pd.read_excel(xls, xls.sheet_names[1]) if len(xls.sheet_names) > 1 else main_df
            for df, name in [(main_df, "主数据"), (aux_df, "辅助数据")]:
                for col in FIELDS:
                    if col not in df.columns:
                        self.show_console()
                        self.console_stream.write(f"[ERROR] {name}缺少列: {col}\n")
                        raise ValueError(f"{name}缺少列: {col}")
            main_df = main_df[FIELDS]
            aux_df = aux_df[FIELDS]
            dataframe_to_table(main_df, self.table_main)
            dataframe_to_table(aux_df, self.table_aux)
            self.label.setText(f"已加载Excel: {file_path}\n主数据{main_df.shape[0]}行，辅助数据{aux_df.shape[0]}行")
            self.show_console()
            self.console_stream.write(f"[INFO] Excel加载成功：{file_path}\n")
        except Exception as e:
            self.show_console()
            self.console_stream.write(f"[ERROR] 读取Excel失败:\n{str(e)}\n")
            self.console_stream.write(traceback.format_exc())

            QMessageBox.critical(self, "错误", f"读取Excel失败:\n{str(e)}")

    def process_and_plot(self):
        import traceback
        main_df = table_to_dataframe(self.table_main)
        aux_df = table_to_dataframe(self.table_aux)
        self.show_console()
        if main_df[VECTOR_FIELDS].dropna().shape[0] == 0 or aux_df[VECTOR_FIELDS].dropna().shape[0] == 0:
            self.label.setText("请填写主数据和辅助数据的所有磁场分量")
            self.console_stream.write("[WARN] 数据不完整，无法计算。\n")
            return

        try:
            main_vec = main_df[VECTOR_FIELDS].dropna().astype(float).values
            aux_vec = aux_df[VECTOR_FIELDS].dropna().astype(float).values
            main_geo = main_df[["经度(°)", "纬度(°)", "垂向深度(m)", "井口海拔(m)"]].dropna().astype(float).values
            aux_geo = aux_df[["经度(°)", "纬度(°)", "垂向深度(m)", "井口海拔(m)"]].dropna().astype(float).values
        except Exception as e:
            self.label.setText(f"数据格式错误: {e}")
            self.console_stream.write(f"[ERROR] 数据格式错误: {e}\n")
            self.console_stream.write(traceback.format_exc())
            return

        # 剔除异常
        if aux_vec.shape[0] == 0:
            self.label.setText("辅助数据为空")
            self.console_stream.write("[WARN] 辅助数据为空。\n")
            return
        if aux_vec.shape[0] == 1:
            aux_valid = aux_vec
            aux_geo_valid = aux_geo
            aux_mean = aux_vec[0]
        else:
            outlier_mask = detect_outliers_zscore(aux_vec)
            aux_valid = aux_vec[~outlier_mask]
            aux_geo_valid = aux_geo[~outlier_mask]
            if aux_valid.shape[0] == 0:
                aux_valid = aux_vec
                aux_geo_valid = aux_geo
            aux_mean = weighted_average(aux_valid)

        self.console_stream.write(f"[INFO] 有效辅助数据点数：{aux_valid.shape[0]}\n")
        self.console_stream.write(f"[INFO] 辅助数据均值：{aux_mean}\n")

        # 修正主数据
        main_corrected = correct_main_data(main_vec, aux_mean)
        if len(main_corrected) > 0:
            self.console_stream.write(f"[INFO] 修正后主数据首行：{main_corrected[0]}\n")
        else:
            self.console_stream.write(f"[WARN] 修正后主数据为空。\n")

        # 可视化
        self.figure.clear()
        ax = self.figure.add_subplot(111, projection='3d')
        if aux_geo_valid.shape[0] > 0:
            ax.quiver(
                aux_geo_valid[:, 0], aux_geo_valid[:, 1], aux_geo_valid[:, 2],
                aux_valid[:, 0], aux_valid[:, 1], aux_valid[:, 2],
                color='green', length=200, normalize=True, label='辅助数据'
            )
        if main_geo.shape[0] > 0:
            ax.quiver(
                main_geo[:, 0], main_geo[:, 1], main_geo[:, 2],
                main_vec[:, 0], main_vec[:, 1], main_vec[:, 2],
                color='blue', length=200, normalize=True, label='主数据'
            )
        if main_geo.shape[0] > 0:
            ax.quiver(
                main_geo[:, 0], main_geo[:, 1], main_geo[:, 2],
                main_corrected[:, 0], main_corrected[:, 1], main_corrected[:, 2],
                color='red', length=200, normalize=True, label='修正后主数据'
            )
        ax.set_xlabel('经度(°)')
        ax.set_ylabel('纬度(°)')
        ax.set_zlabel('垂向深度(m)')
        ax.set_title('地磁场向量修正可视化\n蓝:主数据 绿:辅助(去异常) 红:修正后主数据')
        self.canvas.draw()
        self.label.setText("修正与可视化完成。")
        self.console_stream.write("[INFO] 可视化已更新。\n")

    def toggle_console(self):
        if self.console_toggle_btn.isChecked():
            self.show_console()
        else:
            self.hide_console()

    def show_console(self):
        if not self.console_shown:
            self.console_text.show()
            # redirect print to our custom stream
            sys.stdout = self.console_stream
            sys.stderr = self.console_stream
            self.console_shown = True
            self.console_toggle_btn.setText("隐藏控制台")

    def hide_console(self):
        if self.console_shown:
            self.console_text.hide()
            # Restore to a dummy sink, so nothing goes to the external console
            sys.stdout = open(os.devnull, 'w')
            sys.stderr = open(os.devnull, 'w')
            self.console_shown = False
            self.console_toggle_btn.setText("显示控制台")

import os

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MagneticCorrectionGeoApp()
    window.show()
    sys.exit(app.exec_())