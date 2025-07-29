import sys
import serial
import serial.tools.list_ports
import numpy as np
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QHBoxLayout, QLabel, QComboBox, QPushButton, 
                            QLineEdit, QTextEdit, QGroupBox, QFormLayout,
                            QMessageBox, QSpinBox, QFileDialog, QCheckBox)  # 添加QFileDialog
from PyQt5.QtCore import QThread, pyqtSignal, Qt
from PyQt5.QtGui import QFont
import matplotlib
matplotlib.use('Qt5Agg')
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import matplotlib.pyplot as plt
from collections import deque
import csv
from datetime import datetime

# 设置matplotlib支持中文显示
plt.rcParams["axes.unicode_minus"] = False  # 解决负号显示问题

# 新增：指定中文字体
try:
    # 尝试加载系统中的中文字体
    font = FontProperties(fname=r"C:\Windows\Fonts\simhei.ttf")  # Windows系统SimHei字体
    plt.rcParams["font.family"] = font.get_name()
except:
    # 如果找不到指定字体，使用matplotlib支持的中文字体
    plt.rcParams["font.family"] = ["SimHei", "WenQuanYi Micro Hei", "Heiti TC"]

class SerialThread(QThread):
    """串口数据接收线程"""
    data_received = pyqtSignal(str)
    connection_status = pyqtSignal(bool)
    
    def __init__(self, port, baudrate=9600, timeout=1):
        super().__init__()
        self.port = port
        self.baudrate = baudrate
        self.timeout = timeout
        self.serial = None
        self.running = False
        
    def run(self):
        try:
            self.serial = serial.Serial(self.port, self.baudrate, timeout=self.timeout)
            self.connection_status.emit(True)
            self.running = True
            while self.running:
                if self.serial.in_waiting:
                    data = self.serial.readline().decode('utf-8', errors='replace').strip()
                    self.data_received.emit(data)
        except Exception as e:
            self.connection_status.emit(False)
            self.data_received.emit(f"串口错误: {str(e)}")
        finally:
            self.close()
    
    def send_data(self, data):
        if self.serial and self.serial.is_open:
            try:
                self.serial.write(data.encode('utf-8'))
                return True
            except Exception as e:
                self.data_received.emit(f"发送错误: {str(e)}")
                return False
        return False
    
    def close(self):
        self.running = False
        if self.serial and self.serial.is_open:
            self.serial.close()


class MplCanvas(FigureCanvas):
    """matplotlib画布"""
    def __init__(self, parent=None, width=5, height=4, dpi=100):
        self.fig = Figure(figsize=(width, height), dpi=dpi)
        self.axes = self.fig.add_subplot(111)
        super(MplCanvas, self).__init__(self.fig)
        self.fig.tight_layout()


class SerialMonitor(QMainWindow):
    """串口监控主窗口"""
    def __init__(self):
        super().__init__()
        self.setWindowTitle("串口数据可视化与调试工具")
        self.setGeometry(100, 100, 1200, 700)
        
        self.serial_thread = None
        self.buffer_size = 1000
        self.data_buffer = deque(maxlen=self.buffer_size)
        self.time_buffer = deque(maxlen=self.buffer_size)
        self.current_time = 0
        
        # CSV相关变量
        self.csv_writer = None
        self.csv_file = None
        self.is_saving = False
        
        self.init_ui()
        self.refresh_ports()
        
    def init_ui(self):
        # 创建中心部件和布局
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)
        
        # 左侧控制面板
        left_panel = QVBoxLayout()
        
        # 串口设置组
        serial_group = QGroupBox("串口设置")
        serial_layout = QFormLayout()
        
        self.port_combo = QComboBox()
        self.baudrate_combo = QComboBox()
        self.baudrate_combo.addItems(["9600", "115200", "38400", "57600", "4800"])
        self.baudrate_combo.setCurrentText("9600")
        
        self.refresh_btn = QPushButton("刷新端口")
        self.refresh_btn.clicked.connect(self.refresh_ports)
        
        self.connect_btn = QPushButton("连接")
        self.connect_btn.clicked.connect(self.toggle_connection)
        
        port_layout = QHBoxLayout()
        port_layout.addWidget(self.port_combo)
        port_layout.addWidget(self.refresh_btn)
        
        serial_layout.addRow("串口号:", port_layout)
        serial_layout.addRow("波特率:", self.baudrate_combo)
        serial_layout.addRow(self.connect_btn)
        
        serial_group.setLayout(serial_layout)
        left_panel.addWidget(serial_group)
        
        # 数据发送组
        send_group = QGroupBox("数据发送")
        send_layout = QVBoxLayout()
        
        self.send_text = QLineEdit()
        self.send_text.setPlaceholderText("输入要发送的数据")
        
        self.send_btn = QPushButton("发送")
        self.send_btn.clicked.connect(self.send_data)
        
        self.send_commands = QComboBox()
        self.send_commands.addItems(["", "开启数据采集", "停止数据采集", "复位设备"])
        self.send_commands.currentTextChanged.connect(self.command_selected)
        
        btn_layout = QHBoxLayout()
        btn_layout.addWidget(self.send_btn)
        btn_layout.addWidget(self.send_commands)
        
        send_layout.addWidget(self.send_text)
        send_layout.addLayout(btn_layout)
        
        send_group.setLayout(send_layout)
        left_panel.addWidget(send_group)
        
        # CSV存储组
        csv_group = QGroupBox("CSV存储")
        csv_layout = QVBoxLayout()
        
        self.csv_path_edit = QLineEdit()
        self.csv_path_edit.setReadOnly(True)
        self.csv_path_edit.setPlaceholderText("未选择CSV文件")
        
        self.select_csv_btn = QPushButton("选择CSV文件")
        self.select_csv_btn.clicked.connect(self.select_csv_file)
        
        self.save_csv_checkbox = QCheckBox("保存数据到CSV")
        self.save_csv_checkbox.stateChanged.connect(self.toggle_saving)
        
        csv_header_layout = QHBoxLayout()
        csv_header_layout.addWidget(QLabel("CSV表头:"))
        self.csv_header_edit = QLineEdit("时间,数值")
        csv_header_layout.addWidget(self.csv_header_edit)
        
        csv_layout.addWidget(self.csv_path_edit)
        csv_layout.addWidget(self.select_csv_btn)
        csv_layout.addLayout(csv_header_layout)
        csv_layout.addWidget(self.save_csv_checkbox)
        
        csv_group.setLayout(csv_layout)
        left_panel.addWidget(csv_group)
        
        # 数据显示组
        display_group = QGroupBox("数据显示")
        display_layout = QVBoxLayout()
        
        self.data_display = QTextEdit()
        self.data_display.setReadOnly(True)
        self.data_display.setMinimumHeight(200)
        self.data_display.setFont(QFont("Courier New", 10))
        
        self.clear_display_btn = QPushButton("清空显示")
        self.clear_display_btn.clicked.connect(self.clear_display)
        
        display_layout.addWidget(self.data_display)
        display_layout.addWidget(self.clear_display_btn)
        
        display_group.setLayout(display_layout)
        left_panel.addWidget(display_group)
        
        # 调试信息组
        debug_group = QGroupBox("调试信息")
        debug_layout = QVBoxLayout()
        
        self.debug_display = QTextEdit()
        self.debug_display.setReadOnly(True)
        self.debug_display.setMinimumHeight(150)
        self.debug_display.setFont(QFont("Courier New", 9))
        
        self.test_plot_btn = QPushButton("测试图表")
        self.test_plot_btn.clicked.connect(self.test_plot)
        
        self.show_buffer_btn = QPushButton("显示缓冲区")
        self.show_buffer_btn.clicked.connect(self.show_buffer)
        
        debug_buttons_layout = QHBoxLayout()
        debug_buttons_layout.addWidget(self.test_plot_btn)
        debug_buttons_layout.addWidget(self.show_buffer_btn)
        
        debug_layout.addWidget(self.debug_display)
        debug_layout.addLayout(debug_buttons_layout)
        
        debug_group.setLayout(debug_layout)
        left_panel.addWidget(debug_group)
        
        left_panel.addStretch()
        main_layout.addLayout(left_panel, 3)
        
        # 右侧图表区域
        right_panel = QVBoxLayout()
        
        # 图表类型选择
        chart_type_layout = QHBoxLayout()
        chart_type_layout.addWidget(QLabel("图表类型:"))
        
        self.chart_type = QComboBox()
        self.chart_type.addItems(["折线图", "柱状图", "散点图"])
        self.chart_type.currentTextChanged.connect(self.update_chart_type)
        
        self.update_interval = QSpinBox()
        self.update_interval.setRange(100, 2000)
        self.update_interval.setValue(500)
        self.update_interval.setSuffix(" ms")
        self.update_interval.valueChanged.connect(self.set_update_interval)
        
        chart_type_layout.addWidget(self.chart_type)
        chart_type_layout.addWidget(QLabel("刷新间隔:"))
        chart_type_layout.addWidget(self.update_interval)
        chart_type_layout.addStretch()
        
        right_panel.addLayout(chart_type_layout)
        
        # 创建图表画布
        self.canvas = MplCanvas(self, width=7, height=5, dpi=100)
        self.canvas.axes.set_title("串口数据可视化")
        self.canvas.axes.set_xlabel("时间")
        self.canvas.axes.set_ylabel("数值")
        right_panel.addWidget(self.canvas)
        
        # 数据解析设置
        parse_group = QGroupBox("数据解析")
        parse_layout = QFormLayout()
        
        self.delimiter_edit = QLineEdit("")
        self.delimiter_edit.setToolTip("数据分隔符，用于解析多值数据")
        
        self.column_edit = QLineEdit("0")
        self.column_edit.setToolTip("要显示的数据列索引，从0开始")
        
        parse_layout.addRow("分隔符:", self.delimiter_edit)
        parse_layout.addRow("数据列:", self.column_edit)
        
        parse_group.setLayout(parse_layout)
        right_panel.addWidget(parse_group)
        
        main_layout.addLayout(right_panel, 7)
        
        # 状态栏
        self.statusBar().showMessage("就绪")
        self.connection_status = QLabel("未连接")
        self.csv_status = QLabel("CSV: 未保存")
        self.statusBar().addPermanentWidget(self.connection_status)
        self.statusBar().addPermanentWidget(self.csv_status)
        
        # 设置定时器更新图表
        self.update_timer = self.startTimer(self.update_interval.value())
        
        # 初始化调试信息
        self.log_debug("程序已启动，请连接串口并接收数据")
    
    def log_debug(self, message):
        """添加调试信息到调试窗口"""
        self.debug_display.append(message)
        self.debug_display.verticalScrollBar().setValue(
            self.debug_display.verticalScrollBar().maximum())
    
    def refresh_ports(self):
        """刷新可用串口列表"""
        self.port_combo.clear()
        ports = serial.tools.list_ports.comports()
        for port in ports:
            self.port_combo.addItem(f"{port.device} - {port.description}")
        self.log_debug(f"已刷新端口列表，找到 {len(ports)} 个串口")
    
    def toggle_connection(self):
        """切换串口连接状态"""
        if self.serial_thread and self.serial_thread.isRunning():
            self.serial_thread.close()
            self.serial_thread.wait()
            self.connect_btn.setText("连接")
            self.connection_status.setText("未连接")
            self.statusBar().showMessage("串口已断开")
            self.log_debug("串口连接已断开")
            # 如果正在保存，停止保存
            if self.is_saving:
                self.toggle_saving(0)
        else:
            try:
                port_name = self.port_combo.currentText().split(" - ")[0]
                baudrate = int(self.baudrate_combo.currentText())
                self.serial_thread = SerialThread(port_name, baudrate)
                self.serial_thread.data_received.connect(self.update_data)
                self.serial_thread.connection_status.connect(self.update_connection_status)
                self.serial_thread.start()
                self.connect_btn.setText("断开")
                self.statusBar().showMessage(f"正在连接到 {port_name}...")
                self.log_debug(f"正在连接到串口: {port_name}，波特率: {baudrate}")
            except Exception as e:
                QMessageBox.critical(self, "连接错误", f"无法连接到串口: {str(e)}")
                self.log_debug(f"串口连接失败: {str(e)}")
    
    def update_connection_status(self, connected):
        """更新连接状态显示"""
        if connected:
            self.connection_status.setText("已连接")
            self.statusBar().showMessage("串口连接成功")
            self.log_debug("串口连接成功")
        else:
            self.connection_status.setText("未连接")
            self.connect_btn.setText("连接")
            self.statusBar().showMessage("串口连接失败")
            self.log_debug("串口连接失败")
    
    def update_data(self, data):
        """更新接收到的数据"""
        self.data_display.append(data)
        self.data_display.verticalScrollBar().setValue(
            self.data_display.verticalScrollBar().maximum())
        
        self.log_debug(f"接收到数据: {data}")
        
        # 解析数据并添加到缓冲区
        try:
            delimiter = self.delimiter_edit.text().strip()  # 去除空格
            column = int(self.column_edit.text())
            
            # 处理空分隔符（数据为单一数值）
            if not delimiter:
                try:
                    # 直接将整个数据转换为数值
                    value = float(data.strip())
                    self.data_buffer.append(value)
                    self.current_time += 1
                    self.time_buffer.append(self.current_time)
                    self.log_debug(f"空分隔符解析成功: 值={value}")
                    
                    # 保存到CSV（如果启用）
                    if self.is_saving and self.csv_writer:
                        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
                        self.csv_writer.writerow([timestamp, value])
                        self.csv_file.flush()
                except ValueError as ve:
                    self.log_debug(f"空分隔符解析失败: 数据'{data}'不是有效数值，错误: {str(ve)}")
                return  # 空分隔符场景处理完毕
            
            # 原有多分隔符解析逻辑（保留）
            if delimiter in data:
                values = data.split(delimiter)
                if column < len(values):
                    try:
                        value = float(values[column].strip())
                        self.data_buffer.append(value)
                        self.current_time += 1
                        self.time_buffer.append(self.current_time)
                        self.log_debug(f"多列解析成功: 列={column}, 值={value}")
                        
                        # 保存到CSV（如果启用）
                        if self.is_saving and self.csv_writer:
                            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
                            self.csv_writer.writerow([timestamp, value])
                            self.csv_file.flush()
                    except ValueError as ve:
                        self.log_debug(f"多列数值转换失败: {str(ve)}")
                else:
                    self.log_debug(f"列索引 {column} 超出范围 (数据列数: {len(values)})")
            else:
                self.log_debug(f"未找到分隔符'{delimiter}'，请检查数据格式或清空分隔符")
        
        except Exception as e:
            self.log_debug(f"数据解析错误: {str(e)}")
    
    def send_data(self):
        """发送数据到串口"""
        data = self.send_text.text().strip()
        if data and self.serial_thread and self.serial_thread.isRunning():
            if self.serial_thread.send_data(data):
                self.data_display.append(f"发送: {data}")
                self.data_display.verticalScrollBar().setValue(
                    self.data_display.verticalScrollBar().maximum())
                self.log_debug(f"发送数据: {data}")
            else:
                self.statusBar().showMessage("发送失败")
                self.log_debug("数据发送失败")
    
    def command_selected(self, command):
        """选择预设命令"""
        if command:
            self.send_text.setText(command)
    
    def clear_display(self):
        """清空数据显示区域"""
        self.data_display.clear()
        self.log_debug("数据显示已清空")
    
    def select_csv_file(self):
        """选择CSV文件"""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "选择CSV文件", "", "CSV Files (*.csv);;All Files (*)")
        
        if file_path:
            # 如果文件不是以.csv结尾，添加.csv扩展名
            if not file_path.lower().endswith('.csv'):
                file_path += '.csv'
            
            self.csv_path_edit.setText(file_path)
            self.log_debug(f"已选择CSV文件: {file_path}")
            
            # 如果正在保存，停止当前保存并重新开始
            if self.is_saving:
                self.toggle_saving(0)
                self.toggle_saving(2)  # 2表示选中状态
    
    def toggle_saving(self, state):
        """切换CSV保存状态"""
        if state == Qt.Checked:  # 2
            csv_path = self.csv_path_edit.text()
            if not csv_path:
                QMessageBox.warning(self, "警告", "请先选择CSV文件")
                self.save_csv_checkbox.setChecked(False)
                self.log_debug("CSV保存失败: 未选择文件")
                return
            
            try:
                # 打开CSV文件
                self.csv_file = open(csv_path, 'w', newline='', encoding='utf-8')
                
                # 写入表头
                headers = self.csv_header_edit.text().split(',')
                self.csv_writer = csv.writer(self.csv_file)
                self.csv_writer.writerow(headers)
                
                self.is_saving = True
                self.csv_status.setText(f"CSV: 正在保存到 {csv_path}")
                self.statusBar().showMessage("开始保存数据到CSV文件")
                self.log_debug(f"已开始保存数据到CSV文件: {csv_path}")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"无法打开CSV文件: {str(e)}")
                self.save_csv_checkbox.setChecked(False)
                self.log_debug(f"CSV文件打开失败: {str(e)}")
        else:  # 0
            if self.csv_file:
                self.csv_file.close()
                self.csv_file = None
                self.csv_writer = None
            
            self.is_saving = False
            self.csv_status.setText("CSV: 未保存")
            self.statusBar().showMessage("已停止保存数据到CSV文件")
            self.log_debug("已停止保存数据到CSV文件")
    
    def update_chart_type(self, chart_type):
        """更新图表类型"""
        self.canvas.axes.clear()
        self.canvas.axes.set_title(f"串口数据可视化 - {chart_type}")
        self.canvas.axes.set_xlabel("时间")
        self.canvas.axes.set_ylabel("数值")
        self.log_debug(f"图表类型已切换为: {chart_type}")
        self.update_chart()
    
    def set_update_interval(self, interval):
        """设置图表更新间隔"""
        self.killTimer(self.update_timer)
        self.update_timer = self.startTimer(interval)
        self.log_debug(f"图表更新间隔已设置为: {interval}ms")
    
    def timerEvent(self, event):
        """定时器事件，更新图表"""
        if event.timerId() == self.update_timer:
            self.update_chart()
    
    def update_chart(self):
        """更新图表显示"""
        if not self.data_buffer:
            self.log_debug("数据缓冲区为空，跳过图表更新")
            return
            
        self.log_debug(f"更新图表: 数据点数量={len(self.data_buffer)}")
        
        self.canvas.axes.clear()
        chart_type = self.chart_type.currentText()
        self.canvas.axes.set_title(f"串口数据可视化 - {chart_type}")
        self.canvas.axes.set_xlabel("时间")
        self.canvas.axes.set_ylabel("数值")
        
        try:
            if chart_type == "折线图":
                self.canvas.axes.plot(self.time_buffer, self.data_buffer, 'b-')
            elif chart_type == "柱状图":
                self.canvas.axes.bar(self.time_buffer, self.data_buffer, width=0.8, color='b')
            elif chart_type == "散点图":
                self.canvas.axes.scatter(self.time_buffer, self.data_buffer, color='b')
            
            self.canvas.axes.grid(True)
            self.canvas.fig.tight_layout()
            self.canvas.draw()
            self.log_debug("图表更新成功")
        except Exception as e:
            self.log_debug(f"图表更新失败: {str(e)}")
    
    def test_plot(self):
        """测试图表功能"""
        self.log_debug("执行图表测试...")
        
        try:
            # 创建简单的测试数据
            x = list(range(1, 11))
            y = [i**2 for i in x]
            
            self.canvas.axes.clear()
            self.canvas.axes.plot(x, y, 'r-', marker='o')
            self.canvas.axes.set_title("测试图表 - 二次函数")
            self.canvas.axes.set_xlabel("X轴")
            self.canvas.axes.set_ylabel("Y轴")
            self.canvas.axes.grid(True)
            self.canvas.fig.tight_layout()
            self.canvas.draw()
            
            self.log_debug("测试图表绘制成功")
        except Exception as e:
            self.log_debug(f"测试图表失败: {str(e)}")
    
    def show_buffer(self):
        """显示当前缓冲区内容"""
        if not self.data_buffer:
            self.log_debug("数据缓冲区为空")
            return
            
        buffer_info = f"数据缓冲区内容 ({len(self.data_buffer)} 个元素):\n"
        buffer_info += f"前5个: {list(self.data_buffer)[:5]}\n"
        buffer_info += f"后5个: {list(self.data_buffer)[-5:]}\n"
        buffer_info += f"最小值: {min(self.data_buffer)}, 最大值: {max(self.data_buffer)}"
        
        self.log_debug(buffer_info)
    
    def closeEvent(self, event):
        """关闭窗口时清理资源"""
        # 停止保存CSV
        if self.is_saving:
            self.toggle_saving(0)
        
        # 断开串口连接
        if self.serial_thread and self.serial_thread.isRunning():
            self.serial_thread.close()
            self.serial_thread.wait()
        
        self.log_debug("程序已关闭")
        event.accept()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = SerialMonitor()
    window.show()
    sys.exit(app.exec_())