"""
串口通信模块

功能：
- 串口扫描（支持检测 Windows 注册表中的虚拟串口）
- 串口数据读取（后台线程）
- 解析下位机发送的温度数据（格式: TEMP:25.5\r\n）
"""

import sys

import serial
import serial.tools.list_ports
from PySide6.QtCore import QThread, Signal


def parse_temperature(data: bytes) -> float | None:
    """
    解析温度数据

    Args:
        data: 串口读取的原始字节数据

    Returns:
        温度值（float），解析失败返回 None

    数据格式: "TEMP:25.5\r\n"
    """
    text = data.decode("ascii", errors="ignore").strip()
    if text.startswith("TEMP:"):
        try:
            return float(text[5:])
        except ValueError:
            return None
    return None


def list_available_ports() -> list[str]:
    """
    扫描系统可用的串口

    Returns:
        串口设备名列表，按 COM 号排序

    扫描方式：
    1. pyserial 标准扫描（检测物理串口、USB 转串口等）
    2. Windows 注册表扫描（检测虚拟串口软件创建的端口，如 com0com、ELTIMA 等）
    """
    # 方法1: pyserial 标准扫描
    ports = [p.device for p in serial.tools.list_ports.comports()]

    # 方法2: Windows 注册表扫描（检测虚拟串口）
    if sys.platform == "win32":
        try:
            import winreg

            key = winreg.OpenKey(
                winreg.HKEY_LOCAL_MACHINE,
                r"HARDWARE\DEVICEMAP\SERIALCOMM",
            )
            i = 0
            while True:
                try:
                    _, value, _ = winreg.EnumValue(key, i)
                    if value not in ports:
                        ports.append(value)
                    i += 1
                except OSError:
                    break
            winreg.CloseKey(key)
        except Exception:
            pass

    # 按 COM 号数字排序（COM1, COM2, COM10, COM11）
    def port_sort_key(port: str):
        if port.startswith("COM"):
            try:
                return (0, int(port[3:]))
            except ValueError:
                pass
        return (1, port)

    ports.sort(key=port_sort_key)
    return ports


class SerialWorker(QThread):
    """
    串口通信工作线程

    信号：
        temperature_received(channel_id, temp): 接收到温度数据
        raw_data_received(channel_id, data): 接收到原始数据
        connection_changed(channel_id, connected): 连接状态变化
        error_occurred(channel_id, error_msg): 发生错误
    """

    temperature_received = Signal(int, float)
    raw_data_received = Signal(int, str)  # (通道号, 原始数据)
    connection_changed = Signal(int, bool)
    error_occurred = Signal(int, str)

    def __init__(self, channel_id: int, port: str, baudrate: int = 9600):
        """
        初始化串口工作线程

        Args:
            channel_id: 通道号（1 或 2）
            port: 串口设备名（如 "COM10"）
            baudrate: 波特率
        """
        super().__init__()
        self.channel_id = channel_id
        self.port = port
        self.baudrate = baudrate
        self._running = False

    def run(self):
        """线程主循环：打开串口，持续读取数据"""
        self._running = True

        # 打开串口
        try:
            ser = serial.Serial(self.port, self.baudrate, timeout=1)
            self.connection_changed.emit(self.channel_id, True)
        except Exception as e:
            self.error_occurred.emit(self.channel_id, f"串口打开失败: {e}")
            self.connection_changed.emit(self.channel_id, False)
            return

        # 循环读取数据
        try:
            while self._running:
                line = ser.readline()
                if not line:
                    continue
                # 发送原始数据
                try:
                    raw_text = line.decode("ascii", errors="ignore").strip()
                    if raw_text:
                        self.raw_data_received.emit(self.channel_id, raw_text)
                except Exception:
                    pass
                # 解析温度
                temp = parse_temperature(line)
                if temp is not None:
                    self.temperature_received.emit(self.channel_id, temp)
        except Exception as e:
            if self._running:
                self.error_occurred.emit(self.channel_id, f"串口读取错误: {e}")
        finally:
            # 关闭串口并通知连接断开
            try:
                ser.close()
            except Exception:
                pass
            self.connection_changed.emit(self.channel_id, False)

    def stop(self):
        """安全停止线程"""
        self._running = False
        self.wait(2000)  # 等待线程结束，超时 2 秒
