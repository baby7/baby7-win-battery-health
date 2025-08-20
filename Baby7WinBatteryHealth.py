import sys
import subprocess
import tempfile
from datetime import timedelta

from PySide6 import QtCore
from PySide6.QtSvg import QSvgRenderer
from PySide6.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QMessageBox, QHBoxLayout, \
    QLabel, QTableWidget, QTableWidgetItem, QHeaderView, QAbstractItemView
from PySide6.QtGui import Qt, QColor, QBrush, QPainter, QPen, QPixmap, QCursor
from PySide6.QtCore import Slot, QPointF, QByteArray, QSize
from lxml import html
from dateutil import parser


def open_url(url):
    # os.system('start ' + str(url))
    subprocess.run('start ' + str(url), shell=True, capture_output=True)


def modify_svg(svg_str: str, scale_factor: float = 1.0) -> QPixmap:
    renderer = QSvgRenderer(QByteArray(svg_str.encode('utf-8')))
    pixmap_size = renderer.defaultSize() * scale_factor
    pixmap = QPixmap(pixmap_size)
    pixmap.fill(QColor(0, 0, 0, 0))
    painter = QPainter(pixmap)
    renderer.render(painter)
    painter.end()
    return pixmap


class BatteryGraphWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.points = []
        self.min_date = None
        self.max_date = None
        self.min_health = 100
        self.max_health = 100
        self.setMinimumHeight(300)

    def set_data(self, data):
        """设置要绘制的数据"""
        self.points = []
        if not data:
            return

        # 转换日期和健康值
        for date_str, _, _, health in data:
            try:
                date = parser.parse(date_str)
                self.points.append((date, health))
            except:
                continue

        if self.points:
            # 计算最小/最大日期和健康值
            self.min_date = min(p[0] for p in self.points)
            self.max_date = max(p[0] for p in self.points)
            self.min_health = min(p[1] for p in self.points)
            self.max_health = max(p[1] for p in self.points)

            # 确保最小健康值不低于0，最大不超过100
            self.min_health = max(0, self.min_health - 5)
            self.max_health = min(100, self.max_health + 5)

        self.update()

    def paintEvent(self, event):
        """绘制电池健康曲线"""
        if not self.points:
            return

        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        # 设置边距
        margin = 40
        width = self.width() - 2 * margin
        height = self.height() - 2 * margin

        # 绘制背景
        painter.fillRect(self.rect(), QColor(240, 240, 240))

        # 绘制坐标轴
        painter.setPen(QPen(Qt.black, 2))
        painter.drawLine(margin, margin, margin, margin + height)  # Y轴
        painter.drawLine(margin, margin + height, margin + width, margin + height)  # X轴

        # 绘制Y轴刻度和标签
        painter.setPen(QPen(Qt.darkGray, 1))
        for i in range(0, 101, 20):
            y = margin + height - (i / 100.0) * height
            painter.drawLine(margin - 5, y, margin + width, y)
            painter.drawText(5, y + 5, f"{i}%")

        # 绘制X轴刻度和标签
        date_range = (self.max_date - self.min_date).days
        if date_range < 1:
            date_range = 1

        # 每隔一定时间点绘制一个刻度
        num_ticks = min(8, date_range)
        interval = max(1, date_range // num_ticks)

        for i in range(0, date_range + 1, interval):
            date = self.min_date + timedelta(days=i)
            x = margin + (i / date_range) * width
            painter.drawLine(x, margin + height, x, margin + height + 5)

            # 格式化日期
            date_str = date.strftime("%Y-%m-%d")
            painter.drawText(x - 20, margin + height + 20, date_str)

        # 绘制曲线
        if len(self.points) > 1:
            path = []
            for i, (date, health) in enumerate(self.points):
                days = (date - self.min_date).days
                x = margin + (days / date_range) * width
                y = margin + height - (health / 100.0) * height

                if i == 0:
                    path.append(QPointF(x, y))
                else:
                    # 使用贝塞尔曲线平滑连接点
                    prev_x, prev_y = path[-1].x(), path[-1].y()
                    ctrl_x1 = prev_x + (x - prev_x) * 0.5
                    ctrl_y1 = prev_y
                    ctrl_x2 = prev_x + (x - prev_x) * 0.5
                    ctrl_y2 = y

                    path.append(QPointF(ctrl_x1, ctrl_y1))
                    path.append(QPointF(ctrl_x2, ctrl_y2))
                    path.append(QPointF(x, y))

            # 绘制曲线
            painter.setPen(QPen(QColor(0, 100, 200), 3))
            painter.drawPolyline(path)

            # 绘制数据点
            painter.setPen(QPen(Qt.darkBlue, 1))
            painter.setBrush(QBrush(QColor(100, 180, 255)))
            for point in path:
                if point in [path[0], path[-1]] or len(path) < 10:  # 只绘制首尾点或当点数少时绘制所有点
                    painter.drawEllipse(point, 5, 5)

        # 绘制标题
        painter.setPen(Qt.darkBlue)
        painter.setFont(self.font())
        painter.drawText(margin, 20, "电池健康曲线")


class BatteryTableWidget(QTableWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setColumnCount(4)
        self.setHorizontalHeaderLabels(["日期", "最大容量 (mWh)", "设计容量 (mWh)", "健康度 (%)"])
        self.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.setAlternatingRowColors(True)

    def set_data(self, data):
        """填充表格数据"""
        self.setRowCount(len(data))

        for row, (date, max_charge, nominal_charge, health) in enumerate(data):
            # 日期
            self.setItem(row, 0, QTableWidgetItem(date))

            # 最大容量
            self.setItem(row, 1, QTableWidgetItem(f"{max_charge:,.0f}"))

            # 设计容量
            self.setItem(row, 2, QTableWidgetItem(f"{nominal_charge:,.0f}"))

            # 健康度
            health_item = QTableWidgetItem(f"{health:.1f}%")

            # 根据健康度设置颜色
            if health > 80:
                health_item.setForeground(QBrush(QColor(0, 128, 0)))  # 绿色
            elif health > 60:
                health_item.setForeground(QBrush(QColor(200, 150, 0)))  # 橙色
            else:
                health_item.setForeground(QBrush(QColor(200, 0, 0)))  # 红色
                health_item.setBackground(QBrush(QColor(255, 230, 230)))  # 浅红色背景

            self.setItem(row, 3, health_item)

        # 按日期排序（最新的在最上面）
        self.sortItems(0, Qt.DescendingOrder)


class MyApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        # 窗口显示后自动生成并加载电池报告
        QApplication.processEvents()
        self.generate_and_load_report()

    def initUI(self):
        self.setWindowTitle('七仔 - Windows电池健康查看工具 V2.0.0')
        layout = QVBoxLayout()
        layout.setSpacing(10)
        layout.setContentsMargins(15, 15, 15, 15)

        top_layout = QHBoxLayout()
        top_layout.setSpacing(10)
        top_layout.setContentsMargins(15, 15, 15, 15)

        # 提示按钮
        report_btn = QPushButton('如何生成html电池健康报告(工作原理)？', self)
        report_btn.setStyleSheet("""
            QPushButton {
                background-color: #0c9ae1;
                padding: 8px;
                font-weight: bold;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #0a80bb;
            }
        """)
        report_btn.setMinimumHeight(40)
        report_btn.clicked.connect(self.show_report_generation_instructions)
        report_btn.setCursor(QCursor(QtCore.Qt.PointingHandCursor))     # 鼠标手形
        top_layout.addWidget(report_btn)

        # Github按钮
        github_btn = QPushButton('', self)
        svg_str = """
<svg xmlns="http://www.w3.org/2000/svg" role="img" viewBox="0 0 24 24" fill="#181717">
    <title>GitHub icon</title>
    <path d="M12 .297c-6.63 0-12 5.373-12 12 0 5.303 3.438 9.8 8.205 11.385.6.113.82-.258.82-.577 0-.285-.01-1.04-.015-2.04-3.338.724-4.042-1.61-4.042-1.61C4.422 18.07 3.633 17.7 3.633 17.7c-1.087-.744.084-.729.084-.729 1.205.084 1.838 1.236 1.838 1.236 1.07 1.835 2.809 1.305 3.495.998.108-.776.417-1.305.76-1.605-2.665-.3-5.466-1.332-5.466-5.93 0-1.31.465-2.38 1.235-3.22-.135-.303-.54-1.523.105-3.176 0 0 1.005-.322 3.3 1.23.96-.267 1.98-.399 3-.405 1.02.006 2.04.138 3 .405 2.28-1.552 3.285-1.23 3.285-1.23.645 1.653.24 2.873.12 3.176.765.84 1.23 1.91 1.23 3.22 0 4.61-2.805 5.625-5.475 5.92.42.36.81 1.096.81 2.22 0 1.606-.015 2.896-.015 3.286 0 .315.21.69.825.57C20.565 22.092 24 17.592 24 12.297c0-6.627-5.373-12-12-12"/>
<style class="stylus">/* 导航条背景透明 */
.tag-container_ksKXH {
    background: none;
}
/* 2020-09-08 添加宽度和缩进 */
.pc-fresh-wrapper-con #container.sam_newgrid #content_left {
    width: 98%
}</style></svg>"""
        github_btn.setIcon(modify_svg(svg_str))
        github_btn.setIconSize(QSize(40, 40))
        github_btn.setStyleSheet("""
            QPushButton {
                background-color: transparent;
                padding: 0px;
            }
        """)
        github_btn.setMaximumWidth(50)
        github_btn.setMinimumHeight(40)
        github_btn.clicked.connect(self.open_github)
        github_btn.setCursor(QCursor(QtCore.Qt.PointingHandCursor))     # 鼠标手形
        top_layout.addWidget(github_btn)

        layout.addLayout(top_layout)

        # 状态标签
        self.status_label = QLabel("正在生成电池报告...")
        self.status_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.status_label)

        # 创建电池健康曲线图
        self.graph_widget = BatteryGraphWidget()
        layout.addWidget(self.graph_widget)

        # 创建数据表格
        self.table_widget = BatteryTableWidget()
        layout.addWidget(self.table_widget)

        # 底部信息
        bottom_layout = QHBoxLayout()
        self.info_label = QLabel("")
        bottom_layout.addWidget(self.info_label)

        refresh_btn = QPushButton('重新生成报告', self)
        refresh_btn.setStyleSheet("padding: 6px;")
        refresh_btn.clicked.connect(self.generate_and_load_report)
        refresh_btn.setCursor(QCursor(QtCore.Qt.PointingHandCursor))     # 鼠标手形
        bottom_layout.addWidget(refresh_btn)

        layout.addLayout(bottom_layout)

        self.setLayout(layout)

        # 设置窗口大小
        self.resize(800, 800)

        # 设置窗口的位置居中
        screen = QApplication.primaryScreen().size()
        self.move((screen.width() - self.width()) // 2,
                  (screen.height() - self.height()) // 2)

    def generate_and_load_report(self):
        """生成电池报告并加载数据"""
        self.status_label.setText("正在生成电池报告...")
        QApplication.processEvents()  # 更新UI

        try:
            # 创建临时文件
            with tempfile.NamedTemporaryFile(suffix='.html', delete=False) as temp_file:
                temp_path = temp_file.name

            # 执行命令生成电池报告
            result = subprocess.run(
                ['powercfg', '/batteryreport', '/output', temp_path],
                capture_output=True,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW
            )

            if result.returncode != 0:
                raise Exception(f"命令执行失败: {result.stderr}")

            # temp_path = "battery-report.html"

            self.status_label.setText(f"报告已生成: {temp_path}")

            # 读取HTML文件内容
            with open(temp_path, 'r', encoding='utf-8') as file:
                html_content = file.read()

            # 解析HTML内容
            data = self.parse_html(html_content)

            if not data:
                self.status_label.setText("未找到电池数据，请检查报告内容")
                return

            # 更新图表和表格
            self.graph_widget.set_data(data)
            self.table_widget.set_data(data)

            # 更新底部信息
            start_date = data[0][0]
            start_health = data[0][3]
            end_date = data[-1][0]
            end_health = data[-1][3]

            self.info_label.setText(
                f"电池健康变化: {start_date} ({start_health:.1f}%) → {end_date} ({end_health:.1f}%) | "
                f"数据点: {len(data)} | 最低健康度: {min(d[3] for d in data):.1f}%"
            )

        except Exception as e:
            self.status_label.setText(f"错误: {str(e)}")
            QMessageBox.critical(self, "错误", f"生成或加载电池报告时出错:\n{str(e)}")
        finally:
            # 删除临时文件
            if os.path.exists(temp_path):
                os.unlink(temp_path)

    def parse_html(self, html_content):
        """解析HTML内容提取电池数据"""
        root = html.fromstring(html_content)
        table = root.xpath('/html/body/table[6]')
        if not table:
            return []

        data = []
        for row in table[0].xpath('.//tr'):
            # 跳过表头
            if row.getparent().tag == 'thead':
                continue

            cols = row.xpath('.//td')
            if len(cols) >= 3 and cols[0] is not None and cols[1] is not None and cols[2] is not None:
                # 提取和清理数据
                date_content = cols[0].text_content().replace('\n', '').replace('\r', '')
                max_charge_content = cols[1].text_content().replace('\n', '').replace('\r', '')
                nominal_charge_content = cols[2].text_content().replace('\n', '').replace('\r', '')

                if "," not in max_charge_content:
                    return []

                date = date_content.strip()[0:10].strip()
                if not date:
                    continue

                max_charge = max_charge_content.strip().replace(',', '').replace(' mWh', '')
                if not max_charge:
                    continue
                max_charge = float(max_charge)

                nominal_charge = nominal_charge_content.strip().replace(',', '').replace(' mWh', '')
                if not nominal_charge:
                    continue
                nominal_charge = float(nominal_charge)

                health = (max_charge / nominal_charge) * 100
                data.append((date, max_charge, nominal_charge, health))

        return data

    @Slot()
    def show_report_generation_instructions(self):
        """显示生成电池健康报告的指令"""
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("生成电池健康报告")
        msg_box.setText("在命令提示符（CMD）中输入以下命令来生成电池健康报告：\n\npowercfg /batteryreport")
        msg_box.setInformativeText("本功能会在窗口打开时自动执行此命令并加载生成的报告")
        msg_box.setStandardButtons(QMessageBox.Ok)
        msg_box.exec_()

    @Slot()
    def open_github(self):
        """打开GitHub项目页面"""
        open_url("https://github.com/baby7/baby7-win-battery-health")


if __name__ == '__main__':
    QApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    app = QApplication(sys.argv)
    ex = MyApp()
    ex.show()
    sys.exit(app.exec())