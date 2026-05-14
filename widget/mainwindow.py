"""
温度采集系统上位机 - 主窗口模块

功能：
- 单路/双路温度采集模式切换
- 串口配置（支持扫描虚拟串口）
- 实时温度显示（大字体）
- pyqtgraph 实时曲线图（独立子图）
- 按通道独立报警（温度超限变红）
- 数据导出（CSV / Excel）
- 配置持久化（config.json）
"""

import json
import os
from collections import deque
from datetime import datetime

import pyqtgraph as pg
from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QAction, QColor, QFont
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QScrollArea,
    QSpinBox,
    QStatusBar,
    QToolBar,
    QVBoxLayout,
    QWidget,
)

from core.alarm import AlarmManager
from core.data_manager import DataManager
from widget.serial_worker import SerialWorker, list_available_ports

# 配置文件路径（项目根目录下的 config.json）
CONFIG_FILE = os.path.join(os.path.dirname(os.path.dirname(__file__)), "config.json")

# 支持的波特率列表
BAUDRATES = ["4800", "9600", "19200", "38400", "57600", "115200"]

# 曲线显示时长选项（标签: 秒数，0 表示全部）
DISPLAY_OPTIONS = {"1分钟": 60, "5分钟": 300, "10分钟": 600, "30分钟": 1800, "全部": 0}

# 各通道曲线颜色
CURVE_COLORS = ["#00BFFF", "#FF6347"]


class MainWindow(QMainWindow):
    """温度采集系统主窗口"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("温度采集系统")
        self.resize(1400, 800)

        # 加载配置
        self.config = self._load_config()

        # 初始化核心模块
        self.data_manager = DataManager(max_records=self.config.get("acquisition", {}).get("max_records", 10000))
        self.alarm_manager = AlarmManager()

        # 串口工作线程字典: {通道号: SerialWorker}
        self.serial_workers: dict[int, SerialWorker] = {}

        # 曲线显示时长（秒）
        self._display_seconds = self.config.get("acquisition", {}).get("display_seconds", 300)

        # UI 控件引用字典（用于后续访问）
        self._combo_widgets: dict[str, dict[int, QWidget]] = {
            "port": {},
            "baudrate": {},
            "alarm_enabled": {},
            "alarm_low": {},
            "alarm_high": {},
        }
        self._temp_labels: dict[int, QLabel] = {}        # 温度显示标签
        self._status_labels: dict[int, QLabel] = {}      # 状态标签
        self._plots: dict[int, pg.PlotWidget] = {}        # 曲线图控件
        self._curves: dict[int, pg.PlotDataItem] = {}     # 曲线数据项
        self._plot_data: dict[int, tuple[list, list]] = {}  # 曲线原始数据 (x列表, y列表)
        self._channel_frames: dict[int, QWidget] = {}    # 通道显示区域容器
        self._log_text: QPlainTextEdit = None             # 原始数据日志控件

        # 初始化 UI
        self._setup_ui()
        self._connect_signals()
        self._apply_config()
        self._scan_ports()

        # 定时刷新曲线（每 200ms）
        self._update_timer = QTimer(self)
        self._update_timer.timeout.connect(self._update_all_curves)
        self._update_timer.start(200)

    # ==================== 配置管理 ====================

    def _load_config(self) -> dict:
        """
        加载配置文件

        Returns:
            配置字典，如果文件不存在或加载失败则返回默认配置
        """
        default = {
            "mode": "single",
            "channels": {
                "1": {
                    "serial": {"port": "COM3", "baudrate": 9600},
                    "alarm": {"enabled": True, "low_limit": 0.0, "high_limit": 50.0},
                },
                "2": {
                    "serial": {"port": "COM4", "baudrate": 9600},
                    "alarm": {"enabled": True, "low_limit": 0.0, "high_limit": 50.0},
                },
            },
            "acquisition": {"max_records": 10000, "display_seconds": 300},
        }
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    saved = json.load(f)
                # 合并保存的配置到默认配置
                for key in default:
                    if key in saved:
                        if isinstance(default[key], dict):
                            default[key].update(saved[key])
                        else:
                            default[key] = saved[key]
            except Exception:
                pass
        return default

    def _save_config(self):
        """保存当前配置到文件"""
        self.config["mode"] = self._mode_combo.currentData()
        # 保存各通道配置
        for ch in [1, 2]:
            ch_key = str(ch)
            if ch_key not in self.config["channels"]:
                self.config["channels"][ch_key] = {
                    "serial": {"port": "COM3", "baudrate": 9600},
                    "alarm": {"enabled": True, "low_limit": 0.0, "high_limit": 50.0},
                }
            if ch in self._combo_widgets["port"]:
                self.config["channels"][ch_key]["serial"]["port"] = self._combo_widgets["port"][ch].currentText()
            if ch in self._combo_widgets["baudrate"]:
                self.config["channels"][ch_key]["serial"]["baudrate"] = int(self._combo_widgets["baudrate"][ch].currentText())
            if ch in self._combo_widgets["alarm_enabled"]:
                self.config["channels"][ch_key]["alarm"]["enabled"] = self._combo_widgets["alarm_enabled"][ch].isChecked()
            if ch in self._combo_widgets["alarm_low"]:
                self.config["channels"][ch_key]["alarm"]["low_limit"] = self._combo_widgets["alarm_low"][ch].value()
            if ch in self._combo_widgets["alarm_high"]:
                self.config["channels"][ch_key]["alarm"]["high_limit"] = self._combo_widgets["alarm_high"][ch].value()
        self.config["acquisition"]["display_seconds"] = self._display_seconds
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(self.config, f, ensure_ascii=False, indent=4)
        except Exception:
            pass

    def _apply_config(self):
        """将保存的配置应用到 UI 控件"""
        mode = self.config.get("mode", "single")
        idx = self._mode_combo.findData(mode)
        if idx >= 0:
            self._mode_combo.setCurrentIndex(idx)
        self._switch_mode(mode)
        # 设置显示时长
        self._display_seconds = self.config.get("acquisition", {}).get("display_seconds", 300)
        for text, secs in DISPLAY_OPTIONS.items():
            if secs == self._display_seconds:
                self._display_duration_combo.setCurrentText(text)
                break

    # ==================== UI 构建 ====================

    def _setup_ui(self):
        """构建主窗口 UI 布局"""
        self._create_toolbar()

        # 中心部件：左侧配置面板 + 右侧数据显示区
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QHBoxLayout(central)

        # 左侧区域：配置面板 + 日志面板
        left_area = QWidget()
        left_layout = QVBoxLayout(left_area)
        left_layout.setContentsMargins(0, 0, 0, 0)

        # 配置面板（可滚动）
        self._config_panel = self._create_config_panel()
        left_layout.addWidget(self._config_panel, 2)  # 比例 2/3

        # 原始数据日志面板
        # log_panel = self._create_log_panel()
        # left_layout.addWidget(log_panel, 1)  # 比例 1/3

        main_layout.addWidget(left_area, 1)  # 左侧整体比例 1

        # 中间数据显示区
        self._data_area = QWidget()
        self._data_layout = QVBoxLayout(self._data_area)
        main_layout.addWidget(self._data_area, 3)  # 比例 3

        # 右侧区域：原始数据日志面板
        self._log_area = QWidget()
        self._log_layout = QVBoxLayout(self._log_area)
        log_panel = self._create_log_panel()
        self._log_layout.addWidget(log_panel)
        main_layout.addWidget(self._log_area, 1)  # 右侧整体比例 1

    

        # 状态栏
        self._status_bar = QStatusBar()
        self.setStatusBar(self._status_bar)
        self._status_bar.showMessage("就绪")

    def _create_toolbar(self):
        """创建工具栏"""
        toolbar = QToolBar("工具栏")
        toolbar.setMovable(False)
        self.addToolBar(toolbar)

        # 创建工具栏按钮
        self._start_all_action = QAction("▶ 全部开始", self)
        self._stop_all_action = QAction("⏹ 全部停止", self)
        self._stop_all_action.setEnabled(False)
        self._export_csv_action = QAction("导出CSV", self)
        self._export_excel_action = QAction("导出Excel", self)
        self._clear_action = QAction("清空数据", self)

        # 添加到工具栏
        toolbar.addAction(self._start_all_action)
        toolbar.addAction(self._stop_all_action)
        toolbar.addSeparator()
        toolbar.addAction(self._export_csv_action)
        toolbar.addAction(self._export_excel_action)
        toolbar.addSeparator()
        toolbar.addAction(self._clear_action)

    def _create_config_panel(self) -> QScrollArea:
        """创建左侧配置面板"""
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setMaximumWidth(280)

        container = QWidget()
        self._config_layout = QVBoxLayout(container)
        self._config_layout.setAlignment(Qt.AlignmentFlag.AlignTop)

        # 模式选择
        mode_group = QGroupBox("模式选择")
        mode_layout = QVBoxLayout(mode_group)
        self._mode_combo = QComboBox()
        self._mode_combo.addItem("单路", "single")
        self._mode_combo.addItem("双路", "dual")
        mode_layout.addWidget(self._mode_combo)
        self._config_layout.addWidget(mode_group)

        # 各通道配置（CH1、CH2）
        self._channel_groups: dict[int, QGroupBox] = {}
        for ch in [1, 2]:
            group = self._create_channel_config(ch)
            self._channel_groups[ch] = group
            self._config_layout.addWidget(group)

        # 采集设置
        acq_group = QGroupBox("采集设置")
        acq_layout = QVBoxLayout(acq_group)

        # 显示时长选择
        dur_layout = QHBoxLayout()
        dur_layout.addWidget(QLabel("显示时长:"))
        self._display_duration_combo = QComboBox()
        for text in DISPLAY_OPTIONS:
            self._display_duration_combo.addItem(text)
        self._display_duration_combo.setCurrentText("5分钟")
        dur_layout.addWidget(self._display_duration_combo)
        acq_layout.addLayout(dur_layout)

        # 最大记录数
        rec_layout = QHBoxLayout()
        rec_layout.addWidget(QLabel("最大记录数:"))
        self._max_records_spin = QSpinBox()
        self._max_records_spin.setRange(1000, 1000000)
        self._max_records_spin.setValue(self.config.get("acquisition", {}).get("max_records", 10000))
        self._max_records_spin.setSingleStep(1000)
        rec_layout.addWidget(self._max_records_spin)
        acq_layout.addLayout(rec_layout)

        self._config_layout.addWidget(acq_group)
        self._config_layout.addStretch()

        scroll.setWidget(container)
        return scroll

    def _create_log_panel(self) -> QGroupBox:
        """创建原始数据日志面板"""
        group = QGroupBox("原始数据日志")
        layout = QVBoxLayout(group)

        # 日志文本框（只读，自动滚动）
        self._log_text = QPlainTextEdit()
        self._log_text.setReadOnly(True)
        self._log_text.setMaximumBlockCount(1000)  # 最多保留 1000 行
        self._log_text.setStyleSheet("font-family: Consolas, monospace; font-size: 12px;")
        self._log_text.setMinimumWidth(250)
        layout.addWidget(self._log_text)

        # 清空按钮
        clear_btn = QPushButton("清空日志")
        clear_btn.clicked.connect(self._log_text.clear)
        layout.addWidget(clear_btn)

        return group

    def _create_channel_config(self, ch: int) -> QGroupBox:
        """
        创建单个通道的配置组

        Args:
            ch: 通道号（1 或 2）

        Returns:
            包含串口设置和报警设置的 QGroupBox
        """
        group = QGroupBox(f"CH{ch} 配置")
        layout = QVBoxLayout(group)

        # 从配置文件读取该通道的默认值
        ch_cfg = self.config.get("channels", {}).get(str(ch), {})
        serial_cfg = ch_cfg.get("serial", {})
        alarm_cfg = ch_cfg.get("alarm", {})

        # ---- 串口设置 ----
        layout.addWidget(QLabel("串口设置"))

        # COM 口选择（可编辑，支持手动输入）
        port_layout = QHBoxLayout()
        port_layout.addWidget(QLabel("COM:"))
        port_combo = QComboBox()
        port_combo.setEditable(True)
        port_combo.setCurrentText(serial_cfg.get("port", f"COM{ch + 2}"))
        port_layout.addWidget(port_combo)
        layout.addLayout(port_layout)
        self._combo_widgets["port"][ch] = port_combo

        # 波特率选择
        baud_layout = QHBoxLayout()
        baud_layout.addWidget(QLabel("波特率:"))
        baud_combo = QComboBox()
        for b in BAUDRATES:
            baud_combo.addItem(b)
        baud_combo.setCurrentText(str(serial_cfg.get("baudrate", 9600)))
        baud_layout.addWidget(baud_combo)
        layout.addLayout(baud_layout)
        self._combo_widgets["baudrate"][ch] = baud_combo

        # 扫描串口按钮
        scan_btn = QPushButton("扫描串口")
        scan_btn.clicked.connect(lambda checked, c=ch: self._scan_ports_for_channel(c))
        layout.addWidget(scan_btn)

        # ---- 报警设置 ----
        layout.addWidget(QLabel("报警设置"))

        # 启用报警复选框
        enabled_cb = QCheckBox("启用报警")
        enabled_cb.setChecked(alarm_cfg.get("enabled", True))
        layout.addWidget(enabled_cb)
        self._combo_widgets["alarm_enabled"][ch] = enabled_cb

        # 温度下限
        low_layout = QHBoxLayout()
        low_layout.addWidget(QLabel("下限:"))
        low_spin = QSpinBox()
        low_spin.setRange(-100, 200)
        low_spin.setValue(int(alarm_cfg.get("low_limit", 0)))
        low_layout.addWidget(low_spin)
        low_layout.addWidget(QLabel("°C"))
        layout.addLayout(low_layout)
        self._combo_widgets["alarm_low"][ch] = low_spin

        # 温度上限
        high_layout = QHBoxLayout()
        high_layout.addWidget(QLabel("上限:"))
        high_spin = QSpinBox()
        high_spin.setRange(-100, 200)
        high_spin.setValue(int(alarm_cfg.get("high_limit", 50)))
        high_layout.addWidget(high_spin)
        high_layout.addWidget(QLabel("°C"))
        layout.addLayout(high_layout)
        self._combo_widgets["alarm_high"][ch] = high_spin

        # 初始化报警管理器的通道配置
        self.alarm_manager.setup_channel(
            ch,
            low=low_spin.value(),
            high=high_spin.value(),
            enabled=enabled_cb.isChecked(),
        )

        return group

    def _build_channel_display(self, ch: int):
        """
        构建单个通道的数据显示区域

        Args:
            ch: 通道号（1 或 2）
        """
        frame = QWidget()
        frame_layout = QVBoxLayout(frame)
        frame_layout.setContentsMargins(0, 0, 0, 0)

        # ---- 头部：温度显示 + 状态 + 独立控制按钮 ----
        header = QWidget()
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(0, 0, 0, 0)

        # 温度大字体显示
        temp_label = QLabel(f"CH{ch}: -- °C")
        temp_label.setFont(QFont("Microsoft YaHei", 24, QFont.Weight.Bold))
        temp_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._temp_labels[ch] = temp_label
        header_layout.addWidget(temp_label)

        # 报警状态标签
        status_label = QLabel("[正常]")
        status_label.setFont(QFont("Microsoft YaHei", 12))
        status_label.setStyleSheet("color: green;")
        self._status_labels[ch] = status_label
        header_layout.addWidget(status_label)

        # 独立开始/停止按钮
        start_btn = QPushButton("▶开始")
        start_btn.setFixedWidth(60)
        start_btn.clicked.connect(lambda checked, c=ch: self._start_channel(c))
        header_layout.addWidget(start_btn)

        stop_btn = QPushButton("⏹停止")
        stop_btn.setFixedWidth(60)
        stop_btn.clicked.connect(lambda checked, c=ch: self._stop_channel(c))
        header_layout.addWidget(stop_btn)

        frame_layout.addWidget(header)

        # ---- 实时曲线图 ----
        plot_widget = pg.PlotWidget(title=f"CH{ch} 实时温度曲线")
        plot_widget.setLabel("left", "温度", units="°C")
        plot_widget.setLabel("bottom", "时间")
        plot_widget.showGrid(x=True, y=True, alpha=0.3)
        plot_widget.setBackground("#1e1e2e")

        # 创建曲线（使用对应通道的颜色）
        color = CURVE_COLORS[ch - 1] if ch <= len(CURVE_COLORS) else "#FFFFFF"
        pen = pg.mkPen(color=color, width=2)
        curve = plot_widget.plot(pen=pen)
        self._curves[ch] = curve
        self._plots[ch] = plot_widget
        self._plot_data[ch] = ([], [])  # 初始化空数据

        frame_layout.addWidget(plot_widget)
        self._data_layout.addWidget(frame)
        self._channel_frames[ch] = frame

    # ==================== 信号连接 ====================

    def _connect_signals(self):
        """连接所有信号槽"""
        # 模式切换
        self._mode_combo.currentIndexChanged.connect(self._on_mode_changed)

        # 工具栏按钮
        self._start_all_action.triggered.connect(self._start_all)
        self._stop_all_action.triggered.connect(self._stop_all)
        self._export_csv_action.triggered.connect(self._export_csv)
        self._export_excel_action.triggered.connect(self._export_excel)
        self._clear_action.triggered.connect(self._clear_data)

        # 采集设置
        self._display_duration_combo.currentTextChanged.connect(self._on_display_duration_changed)
        self._max_records_spin.valueChanged.connect(self._on_max_records_changed)

        # 各通道配置控件
        for ch in [1, 2]:
            # 报警启用/禁用
            self._combo_widgets["alarm_enabled"][ch].toggled.connect(
                lambda checked, c=ch: self.alarm_manager.set_enabled(c, checked)
            )
            # 报警阈值变化
            self._combo_widgets["alarm_low"][ch].valueChanged.connect(
                lambda val, c=ch: self.alarm_manager.set_limits(
                    c, val, self._combo_widgets["alarm_high"][c].value()
                )
            )
            self._combo_widgets["alarm_high"][ch].valueChanged.connect(
                lambda val, c=ch: self.alarm_manager.set_limits(
                    c, self._combo_widgets["alarm_low"][c].value(), val
                )
            )
            # 串口选择变化（触发端口过滤）
            self._combo_widgets["port"][ch].currentTextChanged.connect(
                lambda text, c=ch: self._on_port_changed(c)
            )

        # 报警信号
        self.alarm_manager.alarm_triggered.connect(self._on_alarm_triggered)
        self.alarm_manager.alarm_cleared.connect(self._on_alarm_cleared)

    # ==================== 模式切换 ====================

    def _on_mode_changed(self):
        """模式切换事件处理"""
        new_mode = self._mode_combo.currentData()
        if not new_mode:
            return
        current_mode = self.config.get("mode", "single")
        if new_mode == current_mode:
            return

        # 如果正在采集，弹出确认对话框
        if self.serial_workers:
            reply = QMessageBox.question(
                self,
                "切换模式",
                "切换模式将停止当前采集并清空数据，是否继续？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if reply != QMessageBox.StandardButton.Yes:
                # 用户取消，恢复原选择
                self._mode_combo.blockSignals(True)
                idx = self._mode_combo.findData(current_mode)
                self._mode_combo.setCurrentIndex(idx)
                self._mode_combo.blockSignals(False)
                return

        # 停止采集、清空数据、切换 UI
        self._stop_all()
        self.data_manager.clear()
        self._switch_mode(new_mode)
        self.config["mode"] = new_mode

    def _switch_mode(self, mode: str):
        """
        切换单路/双路模式，重建数据显示区

        Args:
            mode: "single" 或 "dual"
        """
        # 清除旧的数据显示区
        for ch in list(self._channel_frames.keys()):
            frame = self._channel_frames.pop(ch)
            self._data_layout.removeWidget(frame)
            frame.deleteLater()
        for ch in list(self._plots.keys()):
            self._plots.pop(ch, None)
            self._curves.pop(ch, None)
            self._plot_data.pop(ch, None)
            self._temp_labels.pop(ch, None)
            self._status_labels.pop(ch, None)

        # 确定当前模式的通道列表
        channels = [1] if mode == "single" else [1, 2]

        # 显示/隐藏配置面板中的通道配置组
        for ch in channels:
            self._channel_groups[ch].setVisible(True)
        for ch in [1, 2]:
            if ch not in channels:
                self._channel_groups[ch].setVisible(False)

        # 构建数据显示区
        for ch in channels:
            self._build_channel_display(ch)

        self._scan_ports()
        self._update_status_bar()

    # ==================== 串口扫描 ====================

    def _scan_ports(self):
        """扫描所有可用串口并更新下拉框"""
        ports = list_available_ports()
        for ch in [1, 2]:
            if ch not in self._combo_widgets["port"]:
                continue
            combo = self._combo_widgets["port"][ch]
            current = combo.currentText()
            combo.blockSignals(True)
            combo.clear()
            combo.addItems(ports)
            if current in ports:
                combo.setCurrentText(current)
            elif ports:
                combo.setCurrentIndex(0)
            else:
                combo.setCurrentText(current)
            combo.blockSignals(False)

    def _scan_ports_for_channel(self, ch: int):
        """
        为指定通道扫描串口（自动过滤另一通道已选的端口）

        Args:
            ch: 通道号
        """
        ports = list_available_ports()
        # 过滤掉另一通道已选择的串口
        other_ch = 2 if ch == 1 else 1
        if other_ch in self._combo_widgets["port"]:
            other_port = self._combo_widgets["port"][other_ch].currentText()
            available = [p for p in ports if p != other_port]
        else:
            available = ports

        combo = self._combo_widgets["port"][ch]
        current = combo.currentText()
        combo.blockSignals(True)
        combo.clear()
        combo.addItems(available)
        if current in available:
            combo.setCurrentText(current)
        elif available:
            combo.setCurrentIndex(0)
        else:
            combo.setCurrentText(current)
        combo.blockSignals(False)

    def _on_port_changed(self, changed_ch: int):
        """串口选择变化时，更新另一通道的可选列表"""
        other_ch = 2 if changed_ch == 1 else 1
        if other_ch in self._combo_widgets["port"]:
            self._scan_ports_for_channel(other_ch)

    # ==================== 采集控制 ====================

    def _start_all(self):
        """开始所有通道的采集"""
        mode = self._mode_combo.currentData()
        channels = [1] if mode == "single" else [1, 2]
        for ch in channels:
            self._start_channel(ch)

    def _stop_all(self):
        """停止所有通道的采集"""
        for ch in list(self.serial_workers.keys()):
            self._stop_channel(ch)

    def _start_channel(self, ch: int):
        """
        启动指定通道的串口采集

        Args:
            ch: 通道号
        """
        # 已在运行则跳过
        if ch in self.serial_workers and self.serial_workers[ch].isRunning():
            return

        # 读取配置
        port = self._combo_widgets["port"][ch].currentText()
        baudrate = int(self._combo_widgets["baudrate"][ch].currentText())

        if not port:
            QMessageBox.warning(self, "警告", f"CH{ch} 未选择串口")
            return

        # 更新报警配置
        self.alarm_manager.setup_channel(
            ch,
            low=self._combo_widgets["alarm_low"][ch].value(),
            high=self._combo_widgets["alarm_high"][ch].value(),
            enabled=self._combo_widgets["alarm_enabled"][ch].isChecked(),
        )

        # 创建并启动工作线程
        worker = SerialWorker(ch, port, baudrate)
        worker.temperature_received.connect(self._on_temperature)
        worker.raw_data_received.connect(self._on_raw_data)
        worker.connection_changed.connect(self._on_connection_changed)
        worker.error_occurred.connect(self._on_serial_error)
        worker.finished.connect(lambda c=ch: self._on_worker_finished(c))
        self.serial_workers[ch] = worker
        worker.start()

        # 禁用配置面板，更新按钮状态
        self._set_panel_enabled(False)
        self._start_all_action.setEnabled(False)
        self._stop_all_action.setEnabled(True)
        self._update_status_bar()

    def _stop_channel(self, ch: int):
        """
        停止指定通道的采集

        Args:
            ch: 通道号
        """
        if ch in self.serial_workers:
            worker = self.serial_workers.pop(ch)
            worker.stop()

        # 如果没有运行中的通道，恢复配置面板
        if not self.serial_workers:
            self._set_panel_enabled(True)
            self._start_all_action.setEnabled(True)
            self._stop_all_action.setEnabled(False)
        self._update_status_bar()

    def _on_worker_finished(self, ch: int):
        """工作线程结束回调"""
        self.serial_workers.pop(ch, None)
        if not self.serial_workers:
            self._set_panel_enabled(True)
            self._start_all_action.setEnabled(True)
            self._stop_all_action.setEnabled(False)
        self._update_status_bar()

    def _set_panel_enabled(self, enabled: bool):
        """
        启用/禁用配置面板

        Args:
            enabled: True 启用，False 禁用（采集中禁用）
        """
        for ch in [1, 2]:
            if ch in self._combo_widgets["port"]:
                self._combo_widgets["port"][ch].setEnabled(enabled)
            if ch in self._combo_widgets["baudrate"]:
                self._combo_widgets["baudrate"][ch].setEnabled(enabled)
            if ch in self._combo_widgets["alarm_enabled"]:
                self._combo_widgets["alarm_enabled"][ch].setEnabled(enabled)
            if ch in self._combo_widgets["alarm_low"]:
                self._combo_widgets["alarm_low"][ch].setEnabled(enabled)
            if ch in self._combo_widgets["alarm_high"]:
                self._combo_widgets["alarm_high"][ch].setEnabled(enabled)
        self._mode_combo.setEnabled(enabled)
        self._max_records_spin.setEnabled(enabled)

    # ==================== 数据处理 ====================

    def _on_raw_data(self, ch: int, data: str):
        """
        显示原始数据到日志面板

        Args:
            ch: 通道号
            data: 原始数据字符串
        """
        timestamp = datetime.now().strftime("%H:%M:%S.%f")[:-3]
        self._log_text.appendPlainText(f"[{timestamp}] CH{ch}: {data}")
        # 自动滚动到底部
        self._log_text.verticalScrollBar().setValue(
            self._log_text.verticalScrollBar().maximum()
        )

    def _on_temperature(self, ch: int, temp: float):
        """
        接收到温度数据的处理

        Args:
            ch: 通道号
            temp: 温度值
        """
        # 更新温度显示
        if ch in self._temp_labels:
            self._temp_labels[ch].setText(f"CH{ch}: {temp:.1f} °C")

        # 记录数据
        now = datetime.now()
        self.data_manager.add_record(ch, temp, now)

        # 更新曲线数据
        if ch in self._plot_data:
            x_data, y_data = self._plot_data[ch]
            x_data.append(now.timestamp())
            y_data.append(temp)

        # 检查报警
        self.alarm_manager.check(ch, temp)

        count = self.data_manager.record_count(ch)
        self._update_status_bar_data_count(ch, count)

    def _update_all_curves(self):
        """定时刷新所有通道的曲线"""
        for ch in self._curves:
            self._update_curve(ch)

    def _update_curve(self, ch: int):
        """
        更新指定通道的曲线显示

        Args:
            ch: 通道号
        """
        if ch not in self._plot_data or ch not in self._curves:
            return

        x_data, y_data = self._plot_data[ch]
        if not x_data:
            return

        # 根据显示时长裁剪数据
        if self._display_seconds > 0:
            cutoff = datetime.now().timestamp() - self._display_seconds
            while x_data and x_data[0] < cutoff:
                x_data.pop(0)
                y_data.pop(0)

        if not x_data:
            return

        # 计算相对时间（秒），避免 X 轴数值过大
        base = x_data[0]
        rel_x = [x - base for x in x_data]
        self._curves[ch].setData(rel_x, y_data)

    # ==================== 连接状态 ====================

    def _on_connection_changed(self, ch: int, connected: bool):
        """
        串口连接状态变化处理

        Args:
            ch: 通道号
            connected: 是否已连接
        """
        if ch in self._status_labels:
            if connected:
                self._status_labels[ch].setText("[已连接]")
                self._status_labels[ch].setStyleSheet("color: green;")
            else:
                self._status_labels[ch].setText("[断开]")
                self._status_labels[ch].setStyleSheet("color: gray;")
        self._update_status_bar()

    def _on_serial_error(self, ch: int, error: str):
        """
        串口错误处理

        Args:
            ch: 通道号
            error: 错误信息
        """
        QMessageBox.warning(self, f"CH{ch} 错误", error)

    # ==================== 报警处理 ====================

    def _on_alarm_triggered(self, ch: int, alarm_type: str, temp: float):
        """
        报警触发处理

        Args:
            ch: 通道号
            alarm_type: "HIGH" 或 "LOW"
            temp: 当前温度
        """
        # 温度标签变红
        if ch in self._temp_labels:
            self._temp_labels[ch].setStyleSheet("color: red; font-weight: bold;")
        # 状态标签显示报警类型
        if ch in self._status_labels:
            type_text = "超上限" if alarm_type == "HIGH" else "低于下限"
            self._status_labels[ch].setText(f"[报警: {type_text}]")
            self._status_labels[ch].setStyleSheet("color: red;")
        self._update_status_bar()

    def _on_alarm_cleared(self, ch: int, alarm_type: str):
        """
        报警清除处理

        Args:
            ch: 通道号
            alarm_type: "HIGH" 或 "LOW"
        """
        # 恢复温度标签样式
        if ch in self._temp_labels:
            self._temp_labels[ch].setStyleSheet("")
        # 恢复状态标签
        if ch in self._status_labels:
            self._status_labels[ch].setText("[正常]")
            self._status_labels[ch].setStyleSheet("color: green;")
        self._update_status_bar()

    # ==================== 状态栏 ====================

    def _update_status_bar(self):
        """更新状态栏显示"""
        parts = []
        mode = self._mode_combo.currentData()
        channels = [1] if mode == "single" else [1, 2]
        for ch in channels:
            if ch in self.serial_workers and self.serial_workers[ch].isRunning():
                port = self._combo_widgets["port"][ch].currentText()
                parts.append(f"CH{ch}:{port} 已连接")
            else:
                parts.append(f"CH{ch} 未连接")
        self._status_bar.showMessage(" | ".join(parts))

    def _update_status_bar_data_count(self, ch: int, count: int):
        """更新状态栏数据量显示（可扩展）"""
        pass

    # ==================== 采集设置 ====================

    def _on_display_duration_changed(self, text: str):
        """曲线显示时长变化"""
        self._display_seconds = DISPLAY_OPTIONS.get(text, 300)

    def _on_max_records_changed(self, value: int):
        """最大记录数变化"""
        self.data_manager.max_records = value
        # 调整现有 deque 的 maxlen
        for ch in self.data_manager._data:
            self.data_manager._data[ch] = deque(self.data_manager._data[ch], maxlen=value)

    # ==================== 数据导出 ====================

    def _export_csv(self):
        """导出 CSV 文件"""
        channels = self._get_export_channels()
        if not channels:
            return
        filepath, _ = QFileDialog.getSaveFileName(
            self, "导出CSV", f"temperature_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            "CSV文件 (*.csv)"
        )
        if filepath:
            try:
                self.data_manager.export_csv(filepath, channels)
                QMessageBox.information(self, "导出成功", f"数据已导出到:\n{filepath}")
            except Exception as e:
                QMessageBox.warning(self, "导出失败", str(e))

    def _export_excel(self):
        """导出 Excel 文件"""
        channels = self._get_export_channels()
        if not channels:
            return
        filepath, _ = QFileDialog.getSaveFileName(
            self, "导出Excel", f"temperature_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            "Excel文件 (*.xlsx)"
        )
        if filepath:
            try:
                self.data_manager.export_excel(filepath, channels)
                QMessageBox.information(self, "导出成功", f"数据已导出到:\n{filepath}")
            except Exception as e:
                QMessageBox.warning(self, "导出失败", str(e))

    def _get_export_channels(self) -> list[int]:
        """
        获取要导出的通道列表（双路模式下弹出选择对话框）

        Returns:
            通道号列表
        """
        mode = self._mode_combo.currentData()
        # 单路模式直接返回 CH1
        if mode == "single":
            return [1]

        # 双路模式：检查哪些通道有数据
        channels = []
        for ch in [1, 2]:
            if self.data_manager.record_count(ch) > 0:
                channels.append(ch)
        if not channels:
            QMessageBox.information(self, "提示", "没有数据可导出")
            return []
        # 只有一个通道有数据，直接返回
        if len(channels) == 1:
            return channels

        # 两个通道都有数据，弹出选择对话框
        items = [f"CH{ch} ({self.data_manager.record_count(ch)}条)" for ch in channels]
        selected, ok = QInputDialog.getItem(
            self, "选择导出通道", "导出通道:", items, 0, False
        )
        if ok and selected:
            ch = int(selected.split("CH")[1].split(" ")[0])
            return [ch]
        return []

    # ==================== 数据清空 ====================

    def _clear_data(self):
        """清空所有数据（带确认对话框）"""
        reply = QMessageBox.question(
            self, "清空数据", "确定要清空所有数据吗？",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if reply == QMessageBox.StandardButton.Yes:
            self.data_manager.clear()
            # 清空曲线数据
            for ch in self._plot_data:
                self._plot_data[ch] = ([], [])
            # 重置温度显示
            for ch in self._temp_labels:
                self._temp_labels[ch].setText(f"CH{ch}: -- °C")
                self._temp_labels[ch].setStyleSheet("")
            # 重置状态标签
            for ch in self._status_labels:
                self._status_labels[ch].setText("[正常]")
                self._status_labels[ch].setStyleSheet("color: green;")

    # ==================== 窗口事件 ====================

    def closeEvent(self, event):
        """窗口关闭事件：停止采集、保存配置"""
        self._stop_all()
        self._save_config()
        event.accept()
