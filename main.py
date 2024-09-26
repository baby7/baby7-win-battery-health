import sys

import dateutil
from PySide2.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QMessageBox, QHBoxLayout, \
    QLabel
from PySide2.QtGui import Qt
from PySide2.QtCore import Slot
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.dates import AutoDateLocator, AutoDateFormatter, date2num
from matplotlib.figure import Figure
from lxml import html
from dateutil import parser

# 设置字体以支持中文
plt.rcParams['font.family'] = ['SimHei']  # 使用黑体字体
plt.rcParams['axes.unicode_minus'] = False  # 正常显示负号


class MplCanvas(FigureCanvas):
    def __init__(self, parent=None, width=5, height=4, dpi=100):
        fig = Figure(figsize=(width, height), dpi=dpi)
        self.axes = fig.add_subplot(111)
        super(MplCanvas, self).__init__(fig)


class MyApp(QWidget):
    canvas = None
    show_label = None

    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('七仔的Windows电池健康曲线查看软件V1.0.0')
        layout = QVBoxLayout()

        # 提示按钮
        report_btn = QPushButton('如何生成html电池健康报告？', self)
        report_btn.clicked.connect(self.show_report_generation_instructions)  # 绑定点击事件
        layout.addWidget(report_btn)

        # 创建按钮
        btn = QPushButton('导入win电池健康html文件', self)
        btn.clicked.connect(self.import_html_file)  # 绑定点击事件
        layout.addWidget(btn)

        # 创建一个用于显示图表的MplCanvas
        self.canvas = MplCanvas(self, width=5, height=4, dpi=100)
        layout.addWidget(self.canvas)

        # 添加一个布局来显示开始和结束的信息
        info_layout = QHBoxLayout()
        self.show_label = QLabel()
        info_layout.addWidget(self.show_label)
        layout.addLayout(info_layout)

        self.setLayout(layout)

        # 获取屏幕尺寸并设置窗口宽度为屏幕宽度的80%
        screen = QApplication.primaryScreen()
        screen_size = screen.size()
        self.setFixedWidth(int(screen_size.width() * 0.8))

        # 设置窗口的位置居中
        self.move((screen_size.width() - self.width()) // 2,
                  (screen_size.height() - self.height()) // 2)

    @Slot()
    def import_html_file(self):
        """当按钮被点击时，打开文件对话框让用户选择HTML文件"""
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(self, "选择HTML文件", "", "HTML Files (*.html);;All Files (*)", options=options)
        if file_name:
            print(f"选择了文件: {file_name}")
            try:
                # 尝试使用不同的编码格式读取文件
                with open(file_name, 'r', encoding='utf-8') as file:
                    html_content = file.read()
                    self.parse_html_and_update_chart(html_content)
            except UnicodeDecodeError:
                try:
                    with open(file_name, 'r', encoding='gbk') as file:
                        html_content = file.read()
                        self.parse_html_and_update_chart(html_content)
                except UnicodeDecodeError:
                    print("无法识别文件编码，请检查文件格式。")

    @Slot()
    def show_report_generation_instructions(self):
        """当按钮被点击时，显示生成电池健康报告的指令"""
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("生成电池健康报告")
        msg_box.setText("在命令提示符（CMD）中输入以下命令来生成电池健康报告：\n\npowercfg /batteryreport")
        msg_box.setStandardButtons(QMessageBox.Ok)
        msg_box.exec_()

    def parse_html_and_update_chart(self, html_content):
        """解析HTML内容并更新图表"""
        root = html.fromstring(html_content)
        table = root.xpath('/html/body/table[6]')
        if not table:
            print("未找到指定的表格，请检查HTML文件。")
            return
        data = []
        for row in table[0].xpath('.//tr'):
            # 检查是否为 <thead> 中的行
            if row.getparent().tag == 'thead':
                continue
            cols = row.xpath('.//td')
            if len(cols) >= 3 and cols[0] is not None and cols[1] is not None and cols[2] is not None:
                # 提取和清理数据
                date_content = cols[0].text_content().replace('\n', '').replace('\r', '')
                max_charge_content = cols[1].text_content().replace('\n', '').replace('\r', '')
                nominal_charge_content = cols[2].text_content().replace('\n', '').replace('\r', '')
                date = date_content.strip()[0:10].strip()
                if not date:
                    continue  # 跳过无效行
                max_charge = max_charge_content.strip().replace(',', '').replace(' mWh', '')
                if not max_charge:
                    continue  # 跳过无效行
                max_charge = float(max_charge)
                nominal_charge = nominal_charge_content.strip().replace(',', '').replace(' mWh', '')
                if not nominal_charge:
                    continue  # 跳过无效行
                nominal_charge = float(nominal_charge)
                health = (max_charge / nominal_charge) * 100
                data.append((date, max_charge, nominal_charge, health))
        self.update_chart(data)

    def update_chart(self, data):
        """更新图表"""
        dates = [d[0] for d in data]
        health_values = [d[3] for d in data]

        # 清除之前的图表
        self.canvas.axes.clear()

        # 将日期字符串转换为日期对象
        dates_parsed = [dateutil.parser.parse(d) for d in dates]
        dates_num = date2num(dates_parsed)

        # 绘制折线图
        self.canvas.axes.plot(dates_num, health_values, marker='o')
        self.canvas.axes.set_xlabel("日期")
        self.canvas.axes.set_ylabel("电池健康值 (%)")

        # 设置日期格式化
        locator = AutoDateLocator()
        formatter = AutoDateFormatter(locator)

        self.canvas.axes.xaxis.set_major_locator(locator)
        self.canvas.axes.xaxis.set_major_formatter(formatter)

        # 设置y轴为百分比形式
        self.canvas.axes.yaxis.set_major_formatter(lambda x, pos: f'{int(x)}%')

        # 更新开始和结束的信息
        start_date = dates[0]
        start_health = health_values[0]
        end_date = dates[-1]
        end_health = health_values[-1]

        self.canvas.axes.set_title("电池健康曲线 开始：{} ({}%)，结束：{} ({}%)"
                                   .format(start_date, start_health, end_date, end_health))

        self.canvas.draw()


if __name__ == '__main__':
    QApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps)
    app = QApplication(sys.argv)
    ex = MyApp()
    ex.show()
    sys.exit(app.exec_())
