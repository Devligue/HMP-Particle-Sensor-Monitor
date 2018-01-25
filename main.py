#!/usr/bin/env python3
# encoding=utf8

# HPM Particle Sensor Monitor

__version__ = '1.0.6'
__author__ = 'K. Dziadowiec <krzysztof.dziadowiec@gmail.com>'

import os
import sys
import winreg
import itertools
import logging
import time
import io
import serial

from binascii import unhexlify, hexlify
from PyQt5 import QtCore, QtGui, QtWidgets

CMDS = {
    'Read Particle Measuring Result': '68010493',
    'Start Particle Measurement': '68010196',
    'Stop Particle Measurement': '68010295',
    'Enable Auto Send': '68014057',
    'Stop Auto Send': '68012077',
}


def initialize_logging(debug):
    logger = logging.getLogger('HPM MONITOR')
    if debug:
        logger.setLevel(logging.DEBUG)
    else:
        logger.setLevel(logging.INFO)
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    stream_handler = logging.StreamHandler()
    stream_handler.setFormatter(formatter)
    logger.addHandler(stream_handler)

    return logger


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
        scb_value = self.connection_toolbar.send_cmd_box.currentText()
        cmd = CMDS[scb_value]

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
            logger.exception('SerialException')
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
        self.ser.write(unhexlify(data.encode('utf-8')))
        logger.debug(
            'WRITE - {}'.format(" ".join(data[i:i + 2] for i in range(0, len(data), 2))))

    def handle_data_read(self, data):
        logger.debug('RECV - {}'.format(' '.join(data)))
        if len(data) == 32:
            # PM2.5
            value = int(data[6], 16) * 256 + int(data[7], 16)
            self.update_pm25_signal.emit(value)
            # PM10
            value1 = int(data[8], 16) * 256 + int(data[9], 16)
            self.update_pm10_signal.emit(value1)

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
            'Read Particle Measuring Result',
            'Start Particle Measurement',
            'Stop Particle Measurement',
            'Enable Auto Send',
            'Stop Auto Send'
        ]
        self.send_cmd_box.addItems(cmds)
        self.addWidget(self.send_cmd_box)

        self.send_cmd_btn = QtWidgets.QPushButton('Send')
        self.addWidget(self.send_cmd_btn)

    def fill_ports_list(self):
        try:
            self.new_ports = list(self.enumerate_serial_ports())
            if self.ports != self.new_ports:
                self.ports = self.new_ports
                self.com_box.clear()
                for portname in self.ports:
                    self.com_box.addItem(portname)

            if not self.com_box.currentText():
                self.com_box.addItem('NONE')
        except:
            pass  # allowing app to run even with no comports available

        QtCore.QTimer.singleShot(1000, self.fill_ports_list)

    def enumerate_serial_ports(self):
        path = 'HARDWARE\\DEVICEMAP\\SERIALCOMM'
        try:
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path)
        except WindowsError:
            return []

        for i in itertools.count():
            try:
                val = winreg.EnumValue(key, i)
                yield str(val[1])
            except EnvironmentError:
                break


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

    logger = initialize_logging(debug=True)

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
