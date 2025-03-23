import sys
import os
import argparse
import pandas as pd
import re
import openpyxl
from openpyxl.styles import PatternFill
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QLabel,
    QFileDialog, QSlider, QHBoxLayout, QLineEdit
)
from PyQt6.QtCore import Qt


# === 时间解析 ===
def find_time(text):
    matches = re.findall(r"\b\d{2}:\d{2}\b(?![^\(]*\))", str(text))
    return matches


class MyTime:
    def __init__(self, time_str):
        self.h, self.m = map(int, time_str.split(":"))

    def __sub__(self, other):
        h1, m1 = self.h, self.m
        h2, m2 = other.h, other.m
        if h1 < h2:
            h1 += 24
        mins = (h1 - h2) * 60 + (m1 - m2)
        return MyTime(f"{mins // 60}:{mins % 60}")

    def __add__(self, other):
        mins = self.h * 60 + self.m + other.h * 60 + other.m
        return MyTime(f"{mins // 60}:{mins % 60}")

    def __str__(self):
        return f"{self.h:02d}:{self.m:02d}"

    def __lt__(self, other):
        return (self.h, self.m) < (other.h, other.m)

    def __gt__(self, other):
        return (self.h, self.m) > (other.h, other.m)


# === 计算逻辑 ===
def calculate_time_diffs(personnes):
    data = {}
    for i in range(len(personnes)):
        items = personnes.iloc[i]
        name = items[0]
        times = []
        for item in items[1:]:
            ts = find_time(item)
            if len(ts) < 2:
                times.append(MyTime('00:00'))
                continue
            time2 = MyTime(ts[-1])
            time1 = MyTime(ts[0])
            times.append(time2 - time1)
        data[name] = times
    return data


def time_check_passed(times, weekdays=3, weekend=8):
    weekdays_idx = [0, 3, 4, 5, 6]
    weekend_idx = [1, 2]

    work_days = sum(1 for idx in weekdays_idx if times[idx] > MyTime("00:30"))
    weekend_hours = sum((times[idx] for idx in weekend_idx), MyTime("00:00"))

    weekend_time = MyTime(f"{weekend}:00")
    return work_days >= weekdays and weekend_hours > weekend_time


def process_xlsx(file_path, weekdays=3, weekend=8):
    df = pd.read_excel(file_path, header=None, usecols=[0, 46, 47, 48, 49, 50, 51, 52, 53])
    personnes = df.iloc[4:]
    data = calculate_time_diffs(personnes)

    wb = openpyxl.load_workbook(file_path)
    ws = wb.active
    red_fill = PatternFill(start_color="FF6666", end_color="FF6666", fill_type="solid")

    for i in range(len(personnes)):
        name = personnes.iloc[i][0]
        if not time_check_passed(data[name], weekdays, weekend):
            for col in range(1, ws.max_column + 1):
                ws.cell(row=i + 5, column=col).fill = red_fill

    output_file = f"output_{os.path.basename(file_path)}"
    wb.save(output_file)
    print(f"处理完成，结果已保存为 {output_file}")


# === PyQt6 GUI ===
class AttendanceApp(QWidget):
    def __init__(self):
        super().__init__()

        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("考勤分析工具")
        self.setGeometry(100, 100, 400, 300)

        layout = QVBoxLayout()

        # 文件选择
        self.file_label = QLabel("拖放或选择Excel文件")
        layout.addWidget(self.file_label)

        self.file_button = QPushButton("选择文件")
        self.file_button.clicked.connect(self.select_file)
        layout.addWidget(self.file_button)

        # 工作日天数滑块
        self.weekdays_label = QLabel("最低工作日天数: 3")
        layout.addWidget(self.weekdays_label)

        self.weekdays_slider = QSlider(Qt.Orientation.Horizontal)
        self.weekdays_slider.setMinimum(1)
        self.weekdays_slider.setMaximum(5)
        self.weekdays_slider.setValue(3)
        self.weekdays_slider.valueChanged.connect(self.update_weekdays_label)
        layout.addWidget(self.weekdays_slider)

        # 周末时长滑块
        self.weekend_label = QLabel("最低周末时长 (小时): 8")
        layout.addWidget(self.weekend_label)

        self.weekend_slider = QSlider(Qt.Orientation.Horizontal)
        self.weekend_slider.setMinimum(1)
        self.weekend_slider.setMaximum(12)
        self.weekend_slider.setValue(8)
        self.weekend_slider.valueChanged.connect(self.update_weekend_label)
        layout.addWidget(self.weekend_slider)

        # 开始处理
        self.process_button = QPushButton("开始处理")
        self.process_button.clicked.connect(self.run_processing)
        layout.addWidget(self.process_button)

        # 退出
        self.quit_button = QPushButton("退出")
        self.quit_button.clicked.connect(self.close)
        layout.addWidget(self.quit_button)

        self.setLayout(layout)

        # 允许拖放文件
        self.setAcceptDrops(True)

    def select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "选择Excel文件", "", "Excel Files (*.xlsx)")
        if file_path:
            self.file_label.setText(file_path)

    def update_weekdays_label(self, value):
        self.weekdays_label.setText(f"最低工作日天数: {value}")

    def update_weekend_label(self, value):
        self.weekend_label.setText(f"最低周末时长 (小时): {value}")

    def run_processing(self):
        file_path = self.file_label.text()
        if not os.path.exists(file_path):
            self.file_label.setText("请先选择Excel文件！")
            return

        weekdays = self.weekdays_slider.value()
        weekend = self.weekend_slider.value()

        process_xlsx(file_path, weekdays, weekend)

        self.file_label.setText("✅ 处理完成！文件已保存")

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        if urls:
            file_path = urls[0].toLocalFile()
            if file_path.endswith(".xlsx"):
                self.file_label.setText(file_path)


# === CLI 入口 ===
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="考勤数据分析")
    parser.add_argument("--cli", action="store_true", help="使用命令行模式")
    args = parser.parse_args()

    if args.cli:
        process_xlsx("testa.xlsx", 3, 8)
    else:
        app = QApplication(sys.argv)
        window = AttendanceApp()
        window.show()
        sys.exit(app.exec())
