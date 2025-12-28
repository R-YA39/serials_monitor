import sys
import serial
import serial.tools.list_ports
import numpy as np
import struct
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QHBoxLayout, QLabel, QComboBox, QPushButton, 
                            QLineEdit, QTextEdit, QGroupBox, QFormLayout,
                            QMessageBox, QSpinBox, QFileDialog, QCheckBox, QGridLayout)
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
import time

# 设置matplotlib支持中文显示
plt.rcParams["axes.unicode_minus"] = False  # 解决负号显示问题

# 新增：指定中文字体
try:
    # 尝试加载系统中的中文字体
    from matplotlib.font_manager import FontProperties
    font = FontProperties(fname=r"C:\Windows\Fonts\simhei.ttf")  # Windows系统SimHei字体
    plt.rcParams["font.family"] = font.get_name()
except:
    # 如果找不到指定字体，使用matplotlib支持的中文字体
    plt.rcParams["font.family"] = ["SimHei", "WenQuanYi Micro Hei", "Heiti TC"]

class SerialThread(QThread):
    """串口数据接收线程"""
    data_received = pyqtSignal(bytes)  # 改为接收原始字节数据
    parsed_data_received = pyqtSignal(list)  # 新增：解析后的数据信号
    connection_status = pyqtSignal(bool)
    
    def __init__(self, port, baudrate=9600, timeout=0.1):
        super().__init__()
        self.port = port
        self.baudrate = baudrate
        self.timeout = timeout
        self.serial = None
        self.running = False
        self.data_buffer = bytearray()  # 用于累积接收的数据
        self.expected_frame_size = 36  # 期望的帧大小
        self.frame_count = 0
        self.last_process_time = 0
        self.min_frame_interval = 0.01  # 最小帧处理间隔，防止溢出(100Hz)
        
    def run(self):
        try:
            self.serial = serial.Serial(
                self.port, 
                baudrate=self.baudrate, 
                timeout=self.timeout,
                bytesize=serial.EIGHTBITS,
                parity=serial.PARITY_NONE,
                stopbits=serial.STOPBITS_ONE
            )
            self.connection_status.emit(True)
            self.running = True
            
            # 增加缓冲区大小以适应高波特率
            self.serial.set_buffer_size(rx_size=16384, tx_size=8192)
            
            while self.running:
                try:
                    # 读取所有可用数据
                    if self.serial.in_waiting:
                        data = self.serial.read(self.serial.in_waiting)
                        if data:
                            self.data_buffer.extend(data)
                            # 尝试解析完整的数据帧
                            self.parse_buffer()
                except Exception as e:
                    self.parsed_data_received.emit([f"读取错误: {str(e)}"])
                    
        except Exception as e:
            self.connection_status.emit(False)
            self.parsed_data_received.emit([f"串口错误: {str(e)}"])
        finally:
            self.close()
    
    def parse_buffer(self):
        """解析数据缓冲区，查找完整的数据帧"""
        # 控制处理频率，防止溢出
        current_time = time.time()
        if current_time - self.last_process_time < self.min_frame_interval:
            return
        self.last_process_time = current_time
        
        buffer_len = len(self.data_buffer)
        
        # 我们需要至少36字节才能尝试解析
        if buffer_len < 36:
            return
        
        # 查找帧尾标记 0x80 0x7F
        for i in range(buffer_len - 1):
            if self.data_buffer[i] == 0x80 and self.data_buffer[i+1] == 0x7F:
                # 找到了可能的帧尾
                frame_end = i + 1
                frame_start = frame_end - 35
                
                if frame_start >= 0:
                    # 提取完整帧
                    if frame_end + 1 <= buffer_len:
                        frame = bytes(self.data_buffer[frame_start:frame_end+1])
                        
                        if len(frame) == 36:
                            self.frame_count += 1
                            self.process_frame(frame)
                            
                            # 从缓冲区中移除已处理的数据
                            del self.data_buffer[:frame_end+1]
                            return
                        else:
                            # 帧长度不正确，清除到当前点
                            del self.data_buffer[:frame_end+1]
                            return
        
        # 如果没有找到完整帧，但缓冲区过大，清空部分数据
        if buffer_len > 1000:
            # 保留最后200字节
            if buffer_len > 200:
                self.data_buffer = self.data_buffer[-200:]
    
    def process_frame(self, frame):
        """处理完整的数据帧"""
        try:
            if len(frame) != 36:
                self.parsed_data_received.emit([f"错误: 帧长度{len(frame)}字节"])
                return
            
            # 验证帧尾标志
            if frame[34] != 0x80 or frame[35] != 0x7F:
                self.parsed_data_received.emit([f"错误: 帧尾{frame[34]:02X} {frame[35]:02X}"])
                return
            
            # 解析前32个字节为8个单精度浮点数
            floats = []
            for i in range(8):
                start_idx = i * 4
                end_idx = start_idx + 4
                try:
                    value = struct.unpack('f', bytes(frame[start_idx:end_idx]))[0]
                    floats.append(value)
                except:
                    floats.append(float('nan'))  # 如果解析失败，用NaN代替
            
            # 解析标志位
            flag1 = frame[32]
            flag2 = frame[33]
            
            # 发射原始字节数据和解析后的数据
            self.data_received.emit(bytes(frame))
            self.parsed_data_received.emit(floats + [flag1, flag2])
                
        except struct.error as e:
            self.parsed_data_received.emit([f"解析错误: {str(e)}"])
        except Exception as e:
            self.parsed_data_received.emit([f"处理错误: {str(e)}"])
    
    def send_data(self, data):
        """发送数据到串口"""
        if self.serial and self.serial.is_open:
            try:
                # 如果数据是字符串形式的十六进制
                if ' ' in data or all(c in "0123456789ABCDEFabcdef" for c in data.replace(' ', '')):
                    # 移除空格并将十六进制字符串转换为字节
                    hex_str = data.replace(' ', '')
                    # 确保十六进制字符串长度为偶数
                    if len(hex_str) % 2 != 0:
                        hex_str = '0' + hex_str
                    data_bytes = bytes.fromhex(hex_str)
                else:
                    # 字符串数据
                    data_bytes = data.encode('utf-8')
                
                self.serial.write(data_bytes)
                return True
            except Exception as e:
                self.parsed_data_received.emit([f"发送错误: {str(e)}"])
                return False
        return False
    
    def send_command(self, cmd_type, bolt_type=None):
        """发送螺栓控制指令"""
        if self.serial and self.serial.is_open:
            try:
                if cmd_type == 0x01:  # 配置指令
                    if bolt_type in [0x04, 0x05, 0x06]:
                        data_bytes = bytes([cmd_type, bolt_type])
                        self.serial.write(data_bytes)
                        return True
                elif cmd_type == 0x02:  # 启动指令
                    data_bytes = bytes([cmd_type])
                    self.serial.write(data_bytes)
                    return True
                return False
            except Exception as e:
                self.parsed_data_received.emit([f"发送指令错误: {str(e)}"])
                return False
        return False
    
    def close(self):
        self.running = False
        if self.serial and self.serial.is_open:
            self.serial.close()
            self.serial = None


class MplCanvas(FigureCanvas):
    """matplotlib画布"""
    def __init__(self, parent=None, width=5, height=4, dpi=100):
        self.fig = Figure(figsize=(width, height), dpi=dpi)
        self.axes = self.fig.add_subplot(111)
        super(MplCanvas, self).__init__(self.fig)
        self.fig.tight_layout()


class BoltRemovalMonitor(QMainWindow):
    """螺栓拆卸执行器上位机软件"""
    def __init__(self):
        super().__init__()
        self.setWindowTitle("螺栓拆卸执行器上位机软件")
        self.setGeometry(100, 100, 1200, 700)
        
        self.serial_thread = None
        self.buffer_size = 500  # 减小缓冲区大小，防止溢出
        # 创建8个数据缓冲区
        self.data_buffers = [deque(maxlen=self.buffer_size) for _ in range(8)]
        self.time_buffer = deque(maxlen=self.buffer_size)
        self.current_time = 0
        
        # CSV相关变量
        self.csv_writer = None
        self.csv_file = None
        self.is_saving = False
        
        # 螺栓状态
        self.bolt_status = "未配置"
        self.bolt_type = None
        
        # 数据统计
        self.total_frames = 0
        self.last_data_time = None
        self.last_chart_update = 0
        self.chart_update_interval = 0.2  # 图表更新间隔(秒)
        
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
        # 添加波特率选项，包括921600和2.5M
        self.baudrate_combo.addItems(["9600", "115200", "38400", "57600", "4800", "19200", "921600", "2500000"])
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
        
        # 螺栓指令控制组
        bolt_control_group = QGroupBox("螺栓拆卸指令控制")
        bolt_control_layout = QVBoxLayout()
        
        # 螺栓配置指令
        config_layout = QVBoxLayout()
        config_label = QLabel("螺栓配置指令:")
        config_label.setFont(QFont("Microsoft YaHei", 10, QFont.Bold))
        config_layout.addWidget(config_label)
        
        # 创建网格布局放置螺栓配置按钮
        config_grid = QGridLayout()
        
        # 螺栓配置按钮
        self.config_m4_btn = QPushButton("配置M4螺栓\n(指令: 01 04)")
        self.config_m4_btn.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; padding: 8px;")
        self.config_m4_btn.clicked.connect(lambda: self.send_bolt_config(0x04))
        self.config_m4_btn.setToolTip("发送配置M4螺栓的指令: 0x01 0x04")
        
        self.config_m5_btn = QPushButton("配置M5螺栓\n(指令: 01 05)")
        self.config_m5_btn.setStyleSheet("background-color: #2196F3; color: white; font-weight: bold; padding: 8px;")
        self.config_m5_btn.clicked.connect(lambda: self.send_bolt_config(0x05))
        self.config_m5_btn.setToolTip("发送配置M5螺栓的指令: 0x01 0x05")
        
        self.config_m6_btn = QPushButton("配置M6螺栓\n(指令: 01 06)")
        self.config_m6_btn.setStyleSheet("background-color: #FF9800; color: white; font-weight: bold; padding: 8px;")
        self.config_m6_btn.clicked.connect(lambda: self.send_bolt_config(0x06))
        self.config_m6_btn.setToolTip("发送配置M6螺栓的指令: 0x01 0x06")
        
        # 启动拆装指令按钮
        self.start_bolt_btn = QPushButton("启动螺栓拆装\n(指令: 02)")
        self.start_bolt_btn.setStyleSheet("background-color: #F44336; color: white; font-weight: bold; padding: 8px;")
        self.start_bolt_btn.clicked.connect(self.send_start_bolt)
        self.start_bolt_btn.setToolTip("发送启动螺栓拆装的指令: 0x02")
        
        # 将按钮添加到网格布局
        config_grid.addWidget(self.config_m4_btn, 0, 0)
        config_grid.addWidget(self.config_m5_btn, 0, 1)
        config_grid.addWidget(self.config_m6_btn, 1, 0)
        config_grid.addWidget(self.start_bolt_btn, 1, 1)
        
        config_layout.addLayout(config_grid)
        bolt_control_layout.addLayout(config_layout)
        
        # 数据统计显示
        stats_layout = QFormLayout()
        
        self.frame_count_label = QLabel("0")
        self.frame_count_label.setStyleSheet("font-weight: bold; color: #2196F3;")
        self.data_rate_label = QLabel("0.0 fps")
        self.data_rate_label.setStyleSheet("font-weight: bold; color: #4CAF50;")
        
        stats_layout.addRow("接收帧数:", self.frame_count_label)
        stats_layout.addRow("数据率:", self.data_rate_label)
        
        bolt_control_layout.addLayout(stats_layout)
        
        # 螺栓状态显示
        bolt_status_layout = QVBoxLayout()
        bolt_status_label = QLabel("螺栓状态:")
        bolt_status_label.setFont(QFont("Microsoft YaHei", 10, QFont.Bold))
        bolt_status_layout.addWidget(bolt_status_label)
        
        self.bolt_status_display = QTextEdit()
        self.bolt_status_display.setReadOnly(True)
        self.bolt_status_display.setMaximumHeight(100)
        self.bolt_status_display.setFont(QFont("Courier New", 9))
        self.bolt_status_display.setPlaceholderText("螺栓状态信息将在这里显示...")
        bolt_status_layout.addWidget(self.bolt_status_display)
        
        bolt_control_layout.addLayout(bolt_status_layout)
        bolt_control_group.setLayout(bolt_control_layout)
        left_panel.addWidget(bolt_control_group)
        
        # 数据发送组
        send_group = QGroupBox("数据发送")
        send_layout = QVBoxLayout()
        
        # 指令格式说明
        format_label = QLabel("指令格式: 01 04(M4) | 01 05(M5) | 01 06(M6) | 02(启动)")
        format_label.setStyleSheet("color: #666; font-size: 10px;")
        send_layout.addWidget(format_label)
        
        self.send_text = QLineEdit()
        self.send_text.setPlaceholderText("输入要发送的数据 (十六进制指令，如: 01 04)")
        
        self.send_btn = QPushButton("发送")
        self.send_btn.clicked.connect(self.send_data)
        
        self.send_commands = QComboBox()
        self.send_commands.addItems(["", "01 04 - 配置M4螺栓", "01 05 - 配置M5螺栓", 
                                    "01 06 - 配置M6螺栓", "02 - 启动拆装"])
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
        self.csv_header_edit = QLineEdit("时间,浮点数1,浮点数2,浮点数3,浮点数4,浮点数5,浮点数6,浮点数7,浮点数8,标志位1,标志位2,螺栓型号")
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
        self.update_interval.setRange(100, 5000)  # 增加更新间隔范围
        self.update_interval.setValue(1000)  # 默认1秒更新一次
        self.update_interval.setSuffix(" ms")
        self.update_interval.valueChanged.connect(self.set_update_interval)
        
        # 数据选择
        self.data_select = QComboBox()
        # 添加8个浮点数选项
        self.data_select.addItems([f"浮点数{i+1}" for i in range(8)])
        self.data_select.addItem("显示前4个")
        self.data_select.addItem("显示后4个")
        self.data_select.currentTextChanged.connect(self.update_chart_data)
        
        chart_type_layout.addWidget(self.chart_type)
        chart_type_layout.addWidget(QLabel("刷新间隔:"))
        chart_type_layout.addWidget(self.update_interval)
        chart_type_layout.addWidget(QLabel("显示数据:"))
        chart_type_layout.addWidget(self.data_select)
        chart_type_layout.addStretch()
        
        right_panel.addLayout(chart_type_layout)
        
        # 创建图表画布
        self.canvas = MplCanvas(self, width=7, height=5, dpi=100)
        self.canvas.axes.set_title("螺栓拆装数据可视化")
        self.canvas.axes.set_xlabel("时间")
        self.canvas.axes.set_ylabel("数值")
        right_panel.addWidget(self.canvas)
        
        # 实时数据显示
        value_group = QGroupBox("实时数据")
        value_layout = QGridLayout()
        
        # 创建8个浮点数标签
        self.float_labels = []
        colors = ['#2196F3', '#4CAF50', '#FF9800', '#9C27B0', 
                  '#00BCD4', '#8BC34A', '#FF5722', '#673AB7']
        
        for i in range(8):
            label = QLabel(f"F{i+1}: 0.000")
            label.setStyleSheet(f"font-weight: bold; font-size: 12px; color: {colors[i]};")
            self.float_labels.append(label)
            value_layout.addWidget(label, i//4, i%4)  # 4列布局
        
        value_group.setLayout(value_layout)
        right_panel.addWidget(value_group)
        
        # 返回消息显示区域（在图表下方添加）
        return_msg_group = QGroupBox("返回消息")
        return_msg_layout = QVBoxLayout()
        
        self.return_msg_display = QTextEdit()
        self.return_msg_display.setReadOnly(True)
        self.return_msg_display.setMaximumHeight(150)
        self.return_msg_display.setFont(QFont("Courier New", 9))
        self.return_msg_display.setPlaceholderText("返回的消息将在这里显示...")
        
        # 添加清空返回消息按钮
        return_msg_btn_layout = QHBoxLayout()
        self.clear_return_msg_btn = QPushButton("清空返回消息")
        self.clear_return_msg_btn.clicked.connect(self.clear_return_messages)
        return_msg_btn_layout.addStretch()
        return_msg_btn_layout.addWidget(self.clear_return_msg_btn)
        
        return_msg_layout.addWidget(self.return_msg_display)
        return_msg_layout.addLayout(return_msg_btn_layout)
        
        return_msg_group.setLayout(return_msg_layout)
        right_panel.addWidget(return_msg_group)
        
        main_layout.addLayout(right_panel, 7)
        
        # 状态栏
        self.statusBar().showMessage("就绪")
        self.connection_status = QLabel("未连接")
        self.csv_status = QLabel("CSV: 未保存")
        self.bolt_status_label = QLabel("螺栓: 未配置")
        self.bolt_type_label = QLabel("型号: 未知")
        self.baudrate_label = QLabel("波特率: 9600")
        self.buffer_status = QLabel("缓冲区: 0")
        self.statusBar().addPermanentWidget(self.connection_status)
        self.statusBar().addPermanentWidget(self.bolt_status_label)
        self.statusBar().addPermanentWidget(self.bolt_type_label)
        self.statusBar().addPermanentWidget(self.baudrate_label)
        self.statusBar().addPermanentWidget(self.buffer_status)
        self.statusBar().addPermanentWidget(self.csv_status)
        
        # 设置定时器更新图表
        self.update_timer = self.startTimer(self.update_interval.value())
        
        # 初始化调试信息
        self.log_debug("螺栓拆卸执行器上位机软件已启动")
        self.log_debug("支持8个浮点数解析")
        
    def send_bolt_config(self, bolt_type):
        """发送螺栓配置指令"""
        if bolt_type == 0x04:
            bolt_name = "M4"
            hex_str = "01 04"
        elif bolt_type == 0x05:
            bolt_name = "M5"
            hex_str = "01 05"
        elif bolt_type == 0x06:
            bolt_name = "M6"
            hex_str = "01 06"
        else:
            return
        
        if self.serial_thread and self.serial_thread.isRunning():
            # 通过串口线程的send_command方法发送
            if self.serial_thread.send_command(0x01, bolt_type):
                # 更新状态显示
                self.bolt_type = bolt_name
                self.bolt_status = "已配置"
                self.bolt_status_label.setText(f"螺栓: {self.bolt_status}")
                self.bolt_type_label.setText(f"型号: {bolt_name}")
                
                # 在状态显示区域记录
                timestamp = datetime.now().strftime('%H:%M:%S')
                self.bolt_status_display.append(f"[{timestamp}] 已发送{bolt_name}螺栓配置指令: {hex_str}")
                
                # 在数据显示区域也显示
                self.data_display.append(f"[指令发送] {hex_str}")
                
                # 在返回消息显示区域显示
                self.return_msg_display.append(f"[{timestamp}] 已发送{bolt_name}螺栓配置指令: {hex_str}")
                self.return_msg_display.verticalScrollBar().setValue(
                    self.return_msg_display.verticalScrollBar().maximum())
                
                self.log_debug(f"已发送{bolt_name}螺栓配置指令: {hex_str}")
                
                # 自动滚动到底部
                self.bolt_status_display.verticalScrollBar().setValue(
                    self.bolt_status_display.verticalScrollBar().maximum())
                self.data_display.verticalScrollBar().setValue(
                    self.data_display.verticalScrollBar().maximum())
            else:
                self.statusBar().showMessage("指令发送失败")
                self.log_debug(f"{bolt_name}螺栓配置指令发送失败: {hex_str}")
        else:
            QMessageBox.warning(self, "连接错误", "串口未连接，请先连接串口")
            self.log_debug("串口未连接，无法发送螺栓配置指令")
    
    def send_start_bolt(self):
        """发送启动螺栓拆装指令"""
        if self.bolt_type is None:
            QMessageBox.warning(self, "操作错误", "请先配置螺栓型号")
            self.log_debug("未配置螺栓型号，无法启动拆装")
            return
        
        if self.serial_thread and self.serial_thread.isRunning():
            # 通过串口线程的send_command方法发送
            if self.serial_thread.send_command(0x02):
                # 更新状态显示
                self.bolt_status = "拆装中"
                self.bolt_status_label.setText(f"螺栓: {self.bolt_status}")
                
                # 在状态显示区域记录
                timestamp = datetime.now().strftime('%H:%M:%S')
                self.bolt_status_display.append(f"[{timestamp}] 已发送启动拆装指令: 02")
                self.bolt_status_display.append(f"[{timestamp}] 开始拆装{self.bolt_type}螺栓...")
                
                # 在数据显示区域也显示
                self.data_display.append(f"[指令发送] 02")
                
                # 在返回消息显示区域显示
                self.return_msg_display.append(f"[{timestamp}] 已发送启动拆装指令: 02")
                self.return_msg_display.append(f"[{timestamp}] 开始拆装{self.bolt_type}螺栓...")
                self.return_msg_display.verticalScrollBar().setValue(
                    self.return_msg_display.verticalScrollBar().maximum())
                
                self.log_debug(f"已发送启动拆装指令: 02")
                
                # 自动滚动到底部
                self.bolt_status_display.verticalScrollBar().setValue(
                    self.bolt_status_display.verticalScrollBar().maximum())
                self.data_display.verticalScrollBar().setValue(
                    self.data_display.verticalScrollBar().maximum())
            else:
                self.statusBar().showMessage("指令发送失败")
                self.log_debug("启动拆装指令发送失败: 02")
        else:
            QMessageBox.warning(self, "连接错误", "串口未连接，请先连接串口")
            self.log_debug("串口未连接，无法发送启动指令")
    
    def log_debug(self, message):
        """添加调试信息到调试窗口"""
        self.debug_display.append(f"[{datetime.now().strftime('%H:%M:%S')}] {message}")
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
            self.bolt_status_label.setText("螺栓: 未连接")
            self.bolt_type_label.setText("型号: 未知")
            self.baudrate_label.setText("波特率: 未连接")
            self.bolt_status = "未连接"
            self.bolt_type = None
            self.statusBar().showMessage("串口已断开")
            self.log_debug("串口连接已断开")
            # 如果正在保存，停止保存
            if self.is_saving:
                self.toggle_saving(0)
        else:
            try:
                port_name = self.port_combo.currentText().split(" - ")[0]
                baudrate_text = self.baudrate_combo.currentText()
                baudrate = int(baudrate_text)
                
                # 更新状态栏显示波特率
                self.baudrate_label.setText(f"波特率: {baudrate_text}")
                
                self.serial_thread = SerialThread(port_name, baudrate)
                self.serial_thread.data_received.connect(self.update_raw_data)
                self.serial_thread.parsed_data_received.connect(self.update_parsed_data)
                self.serial_thread.connection_status.connect(self.update_connection_status)
                self.serial_thread.start()
                self.connect_btn.setText("断开")
                self.statusBar().showMessage(f"正在连接到 {port_name}，波特率: {baudrate_text}...")
                self.log_debug(f"正在连接到串口: {port_name}，波特率: {baudrate_text}")
            except Exception as e:
                QMessageBox.critical(self, "连接错误", f"无法连接到串口: {str(e)}")
                self.log_debug(f"串口连接失败: {str(e)}")
    
    def update_connection_status(self, connected):
        """更新连接状态显示"""
        if connected:
            baudrate_text = self.baudrate_combo.currentText()
            self.connection_status.setText("已连接")
            self.bolt_status_label.setText("螺栓: 已连接")
            self.baudrate_label.setText(f"波特率: {baudrate_text}")
            self.statusBar().showMessage("串口连接成功")
            self.log_debug("串口连接成功")
        else:
            self.connection_status.setText("未连接")
            self.bolt_status_label.setText("螺栓: 未连接")
            self.bolt_type_label.setText("型号: 未知")
            self.baudrate_label.setText("波特率: 未连接")
            self.connect_btn.setText("连接")
            self.statusBar().showMessage("串口连接失败")
            self.log_debug("串口连接失败")
    
    def update_raw_data(self, data_bytes):
        """更新接收到的原始字节数据"""
        # 控制显示频率，防止溢出
        current_time = time.time()
        if hasattr(self, 'last_raw_display') and current_time - self.last_raw_display < 0.1:
            return
        self.last_raw_display = current_time
        
        hex_str = ' '.join([f'{b:02X}' for b in data_bytes[:20]])  # 只显示前20字节
        if len(data_bytes) > 20:
            hex_str += " ..."
        
        # 在返回消息显示区域显示原始数据
        timestamp = datetime.now().strftime('%H:%M:%S.%f')[:-3]
        return_msg = f"[{timestamp}] 接收{len(data_bytes)}字节: {hex_str}"
        self.return_msg_display.append(return_msg)
        self.return_msg_display.verticalScrollBar().setValue(
            self.return_msg_display.verticalScrollBar().maximum())
        
        self.data_display.append(f"[原始数据] {hex_str}")
        self.data_display.verticalScrollBar().setValue(
            self.data_display.verticalScrollBar().maximum())
        
        self.log_debug(f"接收到原始数据: {len(data_bytes)}字节")
        
        # 显示数据包结构
        if len(data_bytes) == 36:
            # 显示数据包解析
            self.log_debug(f"数据包结构: {len(data_bytes)}字节")
            
            # 显示浮点数部分
            try:
                for i in range(8):
                    start_idx = i * 4
                    end_idx = start_idx + 4
                    float_val = struct.unpack('f', bytes(data_bytes[start_idx:end_idx]))[0]
                    self.log_debug(f"  浮点数{i+1}: {float_val:.6f} (字节{start_idx}-{end_idx-1})")
            except:
                self.log_debug(f"  浮点数解析错误")
            
            # 显示标志位
            flag1 = data_bytes[32]
            flag2 = data_bytes[33]
            self.log_debug(f"  标志位1: {flag1:02X} (字节32)")
            self.log_debug(f"  标志位2: {flag2:02X} (字节33)")
            
            # 显示帧尾
            frame_end = ' '.join([f'{b:02X}' for b in data_bytes[34:36]])
            self.log_debug(f"  帧尾(34-35): {frame_end}")
    
    def update_parsed_data(self, data_list):
        """更新解析后的数据"""
        if not data_list:
            return
            
        # 如果数据是错误信息
        if isinstance(data_list[0], str):
            error_msg = data_list[0]
            self.data_display.append(f"[错误] {error_msg}")
            self.return_msg_display.append(f"[错误] {error_msg}")
            self.return_msg_display.verticalScrollBar().setValue(
                self.return_msg_display.verticalScrollBar().maximum())
            self.log_debug(error_msg)
            return
        
        # 解析浮点数和标志位
        if len(data_list) >= 10:  # 8个浮点数 + 2个标志位
            floats = data_list[:8]
            flag1, flag2 = data_list[8], data_list[9]
            
            # 更新数据统计
            self.total_frames += 1
            self.frame_count_label.setText(str(self.total_frames))
            
            # 计算数据率
            current_time = datetime.now()
            if self.last_data_time:
                time_diff = (current_time - self.last_data_time).total_seconds()
                if time_diff > 0:
                    data_rate = 1.0 / time_diff
                    self.data_rate_label.setText(f"{data_rate:.1f} fps")
            self.last_data_time = current_time
            
            # 更新实时数据显示
            for i in range(8):
                if i < len(self.float_labels):
                    self.float_labels[i].setText(f"F{i+1}: {floats[i]:.3f}")
            
            # 添加到数据缓冲区
            for i in range(8):
                if i < len(self.data_buffers):
                    self.data_buffers[i].append(floats[i])
            
            self.current_time += 1
            self.time_buffer.append(self.current_time)
            
            # 更新缓冲区状态
            buffer_len = len(self.data_buffers[0]) if self.data_buffers[0] else 0
            self.buffer_status.setText(f"缓冲区: {self.buffer_size}")
            
            # 控制显示频率
            if self.total_frames % 10 == 0:  # 每10帧显示一次
                # 在数据显示区域显示
                display_str = f"[{datetime.now().strftime('%H:%M:%S.%f')[:-3]}] "
                for i in range(min(4, len(floats))):  # 只显示前4个
                    display_str += f"F{i+1}: {floats[i]:.3f} "
                if len(floats) > 4:
                    display_str += "..."
                
                self.data_display.append(display_str)
                self.data_display.verticalScrollBar().setValue(
                    self.data_display.verticalScrollBar().maximum())
            
            # 在返回消息显示区域显示
            if self.total_frames % 20 == 0:  # 每20帧显示一次
                return_msg = f"[{datetime.now().strftime('%H:%M:%S.%f')[:-3]}] "
                for i in range(min(4, len(floats))):  # 只显示前4个
                    return_msg += f"F{i+1}: {floats[i]:.3f} "
                if len(floats) > 4:
                    return_msg += "..."
                
                self.return_msg_display.append(return_msg)
                self.return_msg_display.verticalScrollBar().setValue(
                    self.return_msg_display.verticalScrollBar().maximum())
            
            # 保存到CSV（如果启用）
            if self.is_saving and self.csv_writer:
                timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
                bolt_type = self.bolt_type if self.bolt_type else "未知"
                row_data = [timestamp] + floats + [f"{flag1:02X}", f"{flag2:02X}", bolt_type]
                self.csv_writer.writerow(row_data)
                if self.total_frames % 10 == 0:  # 每10帧刷新一次文件
                    self.csv_file.flush()
    
    def send_data(self):
        """发送数据到串口"""
        data = self.send_text.text().strip()
        if data and self.serial_thread and self.serial_thread.isRunning():
            if self.serial_thread.send_data(data):
                timestamp = datetime.now().strftime('%H:%M:%S')
                
                # 在数据显示区域显示
                self.data_display.append(f"[{timestamp}] 发送: {data}")
                self.data_display.verticalScrollBar().setValue(
                    self.data_display.verticalScrollBar().maximum())
                
                # 在返回消息显示区域显示
                self.return_msg_display.append(f"[{timestamp}] 发送: {data}")
                self.return_msg_display.verticalScrollBar().setValue(
                    self.return_msg_display.verticalScrollBar().maximum())
                
                self.log_debug(f"发送数据: {data}")
            else:
                self.statusBar().showMessage("发送失败")
                self.log_debug("数据发送失败")
    
    def command_selected(self, command):
        """选择预设命令"""
        if command:
            # 提取指令部分（去除描述文字）
            if " - " in command:
                cmd = command.split(" - ")[0]
            else:
                cmd = command
            self.send_text.setText(cmd)
    
    def clear_display(self):
        """清空数据显示区域"""
        self.data_display.clear()
        self.log_debug("数据显示已清空")
    
    def clear_return_messages(self):
        """清空返回消息显示区域"""
        self.return_msg_display.clear()
        self.log_debug("返回消息显示已清空")
    
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
                self.csv_status.setText(f"CSV: 正在保存")
                self.statusBar().showMessage(f"开始保存数据到CSV文件: {csv_path}")
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
        self.canvas.axes.set_title(f"螺栓拆装数据可视化 - {chart_type}")
        self.canvas.axes.set_xlabel("时间")
        self.canvas.axes.set_ylabel("数值")
        self.log_debug(f"图表类型已切换为: {chart_type}")
        self.update_chart()
    
    def update_chart_data(self, data_type):
        """更新图表显示的数据类型"""
        self.log_debug(f"图表显示数据已切换为: {data_type}")
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
        # 控制图表更新频率
        current_time = time.time()
        if current_time - self.last_chart_update < self.chart_update_interval:
            return
        self.last_chart_update = current_time
        
        if not self.data_buffers[0]:
            return
            
        self.canvas.axes.clear()
        chart_type = self.chart_type.currentText()
        data_type = self.data_select.currentText()
        bolt_type = self.bolt_type if self.bolt_type else "未知"
        
        self.canvas.axes.set_title(f"螺栓拆装数据可视化 - {bolt_type}螺栓 - {chart_type}")
        self.canvas.axes.set_xlabel("时间")
        self.canvas.axes.set_ylabel("数值")
        
        try:
            colors = ['#2196F3', '#4CAF50', '#FF9800', '#9C27B0', 
                      '#00BCD4', '#8BC34A', '#FF5722', '#673AB7']
            
            if data_type.startswith("浮点数"):
                # 显示单个浮点数
                try:
                    idx = int(data_type[3]) - 1
                    if 0 <= idx < len(self.data_buffers):
                        if chart_type == "折线图":
                            self.canvas.axes.plot(self.time_buffer, self.data_buffers[idx], 
                                                '-', color=colors[idx], label=f'浮点数{idx+1}')
                        elif chart_type == "柱状图":
                            self.canvas.axes.bar(self.time_buffer, self.data_buffers[idx], 
                                               width=0.8, color=colors[idx], label=f'浮点数{idx+1}')
                        elif chart_type == "散点图":
                            self.canvas.axes.scatter(self.time_buffer, self.data_buffers[idx], 
                                                   color=colors[idx], label=f'浮点数{idx+1}')
                        self.canvas.axes.legend()
                except:
                    pass
                    
            elif data_type == "显示前4个":
                # 显示前4个浮点数
                for i in range(min(4, len(self.data_buffers))):
                    if chart_type == "折线图":
                        self.canvas.axes.plot(self.time_buffer, self.data_buffers[i], 
                                            '-', color=colors[i], label=f'浮点数{i+1}')
                    elif chart_type == "柱状图":
                        # 柱状图显示多个数据较复杂，这里用折线图替代
                        self.canvas.axes.plot(self.time_buffer, self.data_buffers[i], 
                                            '-', color=colors[i], label=f'浮点数{i+1}')
                    elif chart_type == "散点图":
                        self.canvas.axes.scatter(self.time_buffer, self.data_buffers[i], 
                                               color=colors[i], label=f'浮点数{i+1}', s=10)
                self.canvas.axes.legend()
                
            elif data_type == "显示后4个":
                # 显示后4个浮点数
                for i in range(4, min(8, len(self.data_buffers))):
                    if chart_type == "折线图":
                        self.canvas.axes.plot(self.time_buffer, self.data_buffers[i], 
                                            '-', color=colors[i], label=f'浮点数{i+1}')
                    elif chart_type == "柱状图":
                        self.canvas.axes.plot(self.time_buffer, self.data_buffers[i], 
                                            '-', color=colors[i], label=f'浮点数{i+1}')
                    elif chart_type == "散点图":
                        self.canvas.axes.scatter(self.time_buffer, self.data_buffers[i], 
                                               color=colors[i], label=f'浮点数{i+1}', s=10)
                self.canvas.axes.legend()
            
            self.canvas.axes.grid(True, alpha=0.3)
            self.canvas.fig.tight_layout()
            self.canvas.draw()
            
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
        if not self.data_buffers[0]:
            self.log_debug("数据缓冲区为空")
            return
            
        buffer_info = f"数据缓冲区内容 ({len(self.data_buffers[0])} 个元素):\n"
        
        for i in range(min(4, len(self.data_buffers))):  # 只显示前4个
            if self.data_buffers[i]:
                buffer_info += f"浮点数{i+1} - 前5个: {list(self.data_buffers[i])[:5]}\n"
                buffer_info += f"浮点数{i+1} - 后5个: {list(self.data_buffers[i])[-5:]}\n"
                buffer_info += f"浮点数{i+1} - 最小值: {min(self.data_buffers[i]):.6f}, 最大值: {max(self.data_buffers[i]):.6f}\n"
        
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
        
        self.log_debug("螺栓拆卸执行器上位机软件已关闭")
        event.accept()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = BoltRemovalMonitor()
    window.show()
    sys.exit(app.exec_())