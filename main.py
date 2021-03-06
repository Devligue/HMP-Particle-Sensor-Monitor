#!/usr/bin/env python3
# encoding=utf8

# HPM Particle Sensor Monitor
import os
import sys
#import winreg
import itertools
import logging
import time
import io
from binascii import hexlify
from enum import Enum

import serial
import serial.tools.list_ports
import colorlog
from PyQt5 import QtCore, QtGui, QtWidgets

__version__ = '1.0.7'
__author__ = 'K. Dziadowiec <krzysztof.dziadowiec@gmail.com>'

logger = logging.getLogger(__name__)

DEBUG = True


class CMD(Enum):
    ReadParticleMeasuringResult = b'\x68\x01\x04\x93'
    StartParticleMeasurement = b'\x68\x01\x01\x96'
    StopParticleMeasurement = b'\x68\x01\x02\x95'
    EnableAutoSend = b'\x68\x01\x40\x57'
    StopAutoSend = b'\x68\x01\x20\x77'


def initialize_logging(debug):
    if debug:
        logger.setLevel(logging.DEBUG)
    else:
        logger.setLevel(logging.INFO)
    colored_formatter = colorlog.ColoredFormatter(
        '{time} - {level} - {msg}'.format(
            time='%(purple)s%(asctime)s%(reset)s',
            level='%(log_color)s%(levelname)s%(reset)s',
            msg='%(message)s'),
        datefmt=None,
        reset=True,
        log_colors={
            'DEBUG':    'cyan',
            'INFO':     'green',
            'WARNING':  'yellow',
            'ERROR':    'red',
            'CRITICAL': 'red,bg_white',
        },
        secondary_log_colors={},
        style='%'
    )
    stream_handler = logging.StreamHandler()
    stream_handler.setFormatter(colored_formatter)
    logger.addHandler(stream_handler)


class MonitorWindow(QtWidgets.QMainWindow):

    def __init__(self):
        super(MonitorWindow, self).__init__()
        self.setFixedSize(450, 170)
        self.setWindowTitle('HPM Particle Sensor Monitor')
        self.setStyleSheet("""
            #pm25, #pm10 {
                font-size: 80px;
            }
            """)

        self.monitor = None

        self.central_widget = QtWidgets.QWidget()
        self.setCentralWidget(self.central_widget)

        self.connection_toolbar = ConnectionToolbar()
        self.addToolBar(QtCore.Qt.TopToolBarArea, self.connection_toolbar)
        self.connection_toolbar.connection_btn.clicked.connect(
            self.start_monitor)

        grid_layout = QtWidgets.QGridLayout()
        self.central_widget.setLayout(grid_layout)

        pm25_label = QtWidgets.QLabel('PM 2.5')
        pm25_label.setAlignment(QtCore.Qt.AlignCenter)
        grid_layout.addWidget(pm25_label, 0, 0)

        self.pm25_value = QtWidgets.QLabel('0')
        self.pm25_value.setObjectName('pm25')
        self.pm25_value.setAlignment(QtCore.Qt.AlignCenter)
        grid_layout.addWidget(self.pm25_value, 1, 0, 6, 1)

        pm10_label = QtWidgets.QLabel('PM 10')
        pm10_label.setAlignment(QtCore.Qt.AlignCenter)
        grid_layout.addWidget(pm10_label, 0, 1)

        self.pm10_value = QtWidgets.QLabel('0')
        self.pm10_value.setObjectName('pm10')
        self.pm10_value.setAlignment(QtCore.Qt.AlignCenter)
        grid_layout.addWidget(self.pm10_value, 1, 1, 6, 1)

    def start_monitor(self):
        self.connection_toolbar.connection_btn.clicked.disconnect(
            self.start_monitor)
        self.connection_toolbar.connection_btn.setText('Disconnect')
        self.connection_toolbar.connection_btn.clicked.connect(
            self.stop_monitor)
        self.connection_toolbar.send_cmd_btn.clicked.connect(
            self.send_cmd)

        port_name = self.connection_toolbar.com_box.currentText()

        self.monitor = Monitor(port_name)
        self.monitor.error.connect(self.stop_monitor)
        self.monitor.update_pm25_signal.connect(self.update_pm25)
        self.monitor.update_pm10_signal.connect(self.update_pm10)
        self.monitor.init()

    def stop_monitor(self):
        self.connection_toolbar.connection_btn.clicked.connect(
            self.start_monitor)
        self.connection_toolbar.connection_btn.setText('Connect')
        try:
            self.monitor.data_collector.alive = False
            self.monitor.data_collector.terminate()
            self.monitor.data_collector.wait()
        except:
            pass
        try:
            self.connection_toolbar.connection_btn.clicked.disconnect(
                self.stop_monitor)
            self.connection_toolbar.send_cmd_btn.clicked.disconnect(
                self.send_cmd)
        except:
            pass

    def send_cmd(self):
        selected = self.connection_toolbar.send_cmd_box.currentText()
        cmd = CMD[selected].value

        self.monitor.write_data(cmd)

    def update_pm25(self, value):
        self.pm25_value.setText(str(value))

    def update_pm10(self, value):
        self.pm10_value.setText(str(value))

    def closeEvent(self, event):
        self.stop_monitor()
        QtCore.QCoreApplication.instance().quit()


class Monitor(QtCore.QObject):
    update_pm25_signal = QtCore.pyqtSignal(int)
    update_pm10_signal = QtCore.pyqtSignal(int)
    error = QtCore.pyqtSignal()

    def __init__(self, port_name):
        super(Monitor, self).__init__()
        self.port_name = port_name
        self.ser = None
        self.data_collector = None

    def init(self):
        try:
            self._establish()
            logger.info('Connection Established')
            logger.debug(str(self.ser))

            self.data_collector = DataCollector(self.ser)
            self.data_collector.data_read.connect(self.handle_data_read)
            self.data_collector.finished.connect(self.close_connection)
            self.data_collector.start()
        except serial.serialutil.SerialException as e:
            logger.error(e)
            self.error.emit()
        except Exception as e:
            logger.exception(e)
            self.error.emit()

    def _establish(self):
        '''Open serial port'''
        self.ser = serial.Serial(
            port=str(self.port_name),
            baudrate=9600,
            parity=serial.PARITY_NONE,
            stopbits=serial.STOPBITS_ONE,
            bytesize=serial.EIGHTBITS,
            timeout=0)

    def write_data(self, data):
        self.ser.write(data)
        data = hexlify(data).decode('utf-8')
        logger.debug('WRITE - {}'.format(
            " ".join(data[i:i + 2] for i in range(0, len(data), 2))))

    def handle_data_read(self, data):
        logger.debug('RECV - {}'.format(' '.join(data)))

        if int(data[0], 16) == 0x96:
            if int(data[1], 16) == 0x96:
                logger.error("sensor response: ERROR")
                return

        if int(data[0], 16) == 0xA5:
            if int(data[1], 16) == 0xA5:
                logger.debug("OK")
                return

        if int(data[0], 16) == 0x40:
            if int(data[1], 16) == 0x05:
                if int(data[2], 16) == 0x04:
                    pm25 = int(data[3], 16) * 256 + int(data[4], 16)
                    self.update_pm25_signal.emit(pm25)

                    pm10 = int(data[5], 16) * 256 + int(data[6], 16)
                    self.update_pm10_signal.emit(pm10)
                    logger.info("new oneshot measure")
                    return

        if int(data[0], 16) == 0x42:
            if int(data[1], 16) == 0x4d:
                pm25 = int(data[6], 16) * 256 + int(data[7], 16)
                self.update_pm25_signal.emit(pm25)

                pm10 = int(data[8], 16) * 256 + int(data[9], 16)
                self.update_pm10_signal.emit(pm10)
                logger.info("new auto measure")
                return

        logger.error("sensor response: UNKNOWN")

    def close_connection(self):
        self.ser.close()
        self.update_pm10_signal.emit(0)
        self.update_pm25_signal.emit(0)
        logger.info('Connection Closed')


class DataCollector(QtCore.QThread):
    data_read = QtCore.pyqtSignal(list)
    alive = True

    def __init__(self, ser):
        super(DataCollector, self).__init__()
        self.ser = ser

    def run(self):
        data_string = ''
        while self.alive:
            time.sleep(0.01)

            try:
                buff = io.BytesIO(b'')

                while self.ser.in_waiting:
                    time.sleep(0.015)
                    out = self.ser.read(1)
                    if out != b'':
                        buff.write(out)
                buff_value = buff.getvalue()

                if buff_value:
                    data_string = hexlify(
                        buff_value).decode('utf-8').upper()
                    data = list(
                        map(''.join, zip(*[iter(data_string)] * 2)))
                    self.data_read.emit(data)
                    data_string = ''
            except Exception as e:
                logger.exception(e)
                return


class ConnectionToolbar(QtWidgets.QToolBar):

    def __init__(self):
        super(ConnectionToolbar, self).__init__()
        self.setFloatable(False)
        self.setMovable(False)
        self.setStyleSheet("""
            #com_label,
            #protocol_label {
                margin: 0 5 0 5;
                }

            QComboBox {
                min-width: 40px;
                height: 20px;
                }
            """)

        self.connection_btn = QtWidgets.QPushButton('Connect')
        self.addWidget(self.connection_btn)

        self.com_box = QtWidgets.QComboBox()
        self.com_box.setStatusTip(u'Wybierz port')
        self.ports = []
        self.fill_ports_list()
        self.addWidget(self.com_box)

        self.send_cmd_box = QtWidgets.QComboBox()
        cmds = [
            CMD.ReadParticleMeasuringResult.name,
            CMD.StartParticleMeasurement.name,
            CMD.StopParticleMeasurement.name,
            CMD.EnableAutoSend.name,
            CMD.StopAutoSend.name,
        ]
        self.send_cmd_box.addItems(cmds)
        self.addWidget(self.send_cmd_box)

        self.send_cmd_btn = QtWidgets.QPushButton('Send')
        self.addWidget(self.send_cmd_btn)

    def fill_ports_list(self):
        current_ports = [self.com_box.itemText(i)
                         for i in range(self.com_box.count())]
        new_ports = self.enumerate_serial_ports()

        for missing in set(current_ports) - set(new_ports):
            self.com_box.removeItem(self.com_box.findText(missing))

        for additional in set(new_ports) - set(current_ports):
            self.com_box.addItem(additional)

        QtCore.QTimer.singleShot(1000, self.fill_ports_list)

    @staticmethod
    def enumerate_serial_ports():
        """Finds list of serial ports on this PC."""
        return [str(port[0]) for port in serial.tools.list_ports.comports()]


def create_dark_palette():
    dark_palette = QtGui.QPalette()
    dark_palette_opts = {
        QtGui.QPalette.Window: QtGui.QColor(16, 16, 24),
        QtGui.QPalette.WindowText: QtGui.QColor(214, 218, 213),
        QtGui.QPalette.Base: QtGui.QColor(1, 1, 11),
        QtGui.QPalette.AlternateBase: QtGui.QColor(37, 38, 43),
        QtGui.QPalette.ToolTipBase: QtGui.QColor(16, 16, 24),
        QtGui.QPalette.ToolTipText: QtGui.QColor(214, 218, 213),
        QtGui.QPalette.Text: QtGui.QColor(214, 218, 213),
        QtGui.QPalette.Button: QtGui.QColor(16, 16, 24),
        QtGui.QPalette.ButtonText: QtGui.QColor(214, 218, 213),
        QtGui.QPalette.BrightText: QtGui.QColor(1, 162, 130),
        QtGui.QPalette.Link: QtGui.QColor(1, 162, 130),
        QtGui.QPalette.Highlight: QtGui.QColor(1, 162, 130),
        QtGui.QPalette.HighlightedText: QtGui.QColor(1, 1, 11),
    }
    for key, val in dark_palette_opts.items():
        dark_palette.setColor(key, val)

    return dark_palette


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)

    initialize_logging(debug=DEBUG)

    app.setStyle('Fusion')
    dark_palette = create_dark_palette()
    app.setPalette(dark_palette)

    app.setStyleSheet("""
        QWidget {
            font-family: 'Lato';
            font-size: 13px;
            }
        """)

    main = MonitorWindow()
    main.show()

    os._exit(app.exec_())
