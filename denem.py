import sys
import csv
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QPushButton, QLineEdit, QLabel, QStackedWidget, QFileDialog, QListWidget, QGridLayout, QFrame, QHBoxLayout, QMenuBar, QMenu, QAction, QTextEdit, QProgressBar, QMessageBox
)
from PyQt5.QtCore import QTimer, Qt, QCoreApplication, QThread, pyqtSignal, QObject
import qasync
import asyncio
import xml.etree.ElementTree as ET
import socket

class SCPIClient:
    def __init__(self, ip, port=5025):
        self.ip = ip
        self.port = port
        self.sock = None
        self.channels = []

    async def connect(self):
        self.sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self.sock.settimeout(5)
        try:
            self.sock.connect((self.ip, self.port))
        except socket.error as e:
            raise ConnectionError(f"Connection error to {self.ip}: {e}")

    async def send_command(self, command):
        if not self.sock:
            await self.connect()
        try:
            self.sock.sendall(f"{command}\n".encode('ascii'))
            return await self.read_response()
        except socket.error as e:
            raise ConnectionError(f"Error sending command to {self.ip}: {e}")

    async def read_response(self):
        try:
            response = self.sock.recv(4096).decode('ascii').strip()
            return response
        except socket.error as e:
            raise ConnectionError(f"Error reading response from {self.ip}: {e}")

    def close(self):
        if self.sock:
            self.sock.close()
            self.sock = None

def read_devices_from_xml(file_path):
    tree = ET.parse(file_path)
    root = tree.getroot()
    devices = []

    for device in root.findall('ElectronicLoads/ElectronicLoad'):
        device_info = {
            'id': device.find('ID').text,
            'ip': device.find('IP').text,
            'port': int(device.find('Port').text),
            'channels': []
        }
        for channel in device.find('Channels').findall('Channel'):
            channel_info = {
                'number': int(channel.find('Number').text),
                'name': channel.find('Name').text,
                'type': channel.find('Type').text,
                'value': float(channel.find('Value').text)
            }
            device_info['channels'].append(channel_info)
        devices.append(device_info)

    for device in root.findall('PowerSupplys/PowerSupply'):
        device_info = {
            'id': device.find('ID').text,
            'ip': device.find('IP').text,
            'port': int(device.find('Port').text),
            'voltage_channels': [],
            'current_channels': []
        }
        for channel in device.find('VoltageValues').findall('Channel'):
            channel_info = {
                'number': int(channel.find('Number').text),
                'value': float(channel.find('Value').text)
            }
            device_info['voltage_channels'].append(channel_info)
        for channel in device.find('CurrentValues').findall('Channel'):
            channel_info = {
                'number': int(channel.find('Number').text),
                'value': float(channel.find('Value').text)
            }
            device_info['current_channels'].append(channel_info)
        devices.append(device_info)

    return devices

def read_commands_from_document(file_path):
    tree = ET.parse(file_path)
    root = tree.getroot()
    commands = [command.text for command in root.findall('command')]
    return commands

class AsyncWorker(QObject):
    progress = pyqtSignal(int)
    result = pyqtSignal(str)
    error = pyqtSignal(str)
    done = pyqtSignal()

    def __init__(self, devices):
        super().__init__()
        self.devices = devices
        self.scpi_clients = []

    async def run(self):
        total_devices = len(self.devices)
        for idx, device in enumerate(self.devices):
            client = SCPIClient(device['ip'], device['port'])
            client.channels = device.get('channels', [])
            try:
                await client.connect()
                response = await client.send_command('*IDN?')
                if response:
                    self.result.emit(f"Cihaz {device['id']} bağlı: {response}")
                    if 'channels' in device:
                        for channel in client.channels:
                            await self.initialize_channel(client, channel)
                    else:
                        for channel in device['voltage_channels']:
                            await self.initialize_voltage_channel(client, channel)
                        for channel in device['current_channels']:
                            await self.initialize_current_channel(client, channel)
                    self.scpi_clients.append(client)
                else:
                    self.result.emit(f"Cihaz {device['id']} bağlanamadı.")
            except ConnectionError as e:
                self.error.emit(str(e))
            
            self.progress.emit((idx + 1) * 100 // total_devices)
        self.done.emit()

    async def initialize_channel(self, client, channel):
        number = channel['number']
        type_ = channel['type']
        value = channel['value']
        
        await client.send_command(f'CHAN {number}')
        await client.send_command(f'FUNC {type_}')
        if type_ == 'RES':
            await client.send_command(f'RES {value}')
        elif type_ == 'CURR':
            await client.send_command(f'CURR {value}')
        elif type_ == 'VOLT':
            await client.send_command(f'VOLT {value}')

    async def initialize_voltage_channel(self, client, channel):
        number = channel['number']
        value = channel['value']
        
        await client.send_command(f'VOLT {number},{value}')

    async def initialize_current_channel(self, client, channel):
        number = channel['number']
        value = channel['value']
        
        await client.send_command(f'CURR {number},{value}')

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Elektriksel Test Donanımı Kontrol Sistemi")
        self.setGeometry(100, 100, 1200, 800)
        self.central_widget = QStackedWidget()
        self.setCentralWidget(self.central_widget)
        self.init_ui()
        self.scpi_clients = []  # SCPIClient örnekleri listesi
        self.logging_timer = None
        self.csv_file = 'log_data.csv'
        self.devices = None

    def init_ui(self):
        self.create_menu_bar()
        
        self.main_screen = self.create_main_screen()
        self.device_settings_screen = self.create_device_settings_screen()
        self.logging_settings_screen = self.create_logging_settings_screen()
        self.manual_command_screen = self.create_manual_command_screen()
        self.about_screen = self.create_about_screen()

        self.central_widget.addWidget(self.main_screen)
        self.central_widget.addWidget(self.device_settings_screen)
        self.central_widget.addWidget(self.logging_settings_screen)
        self.central_widget.addWidget(self.manual_command_screen)
        self.central_widget.addWidget(self.about_screen)

        self.central_widget.setCurrentWidget(self.main_screen)

        self.create_footer()

    def create_menu_bar(self):
        menu_bar = self.menuBar()

        file_menu = menu_bar.addMenu("File")
        help_menu = menu_bar.addMenu("Help")

        open_action = QAction("Open", self)
        open_action.triggered.connect(lambda: self.central_widget.setCurrentWidget(self.device_settings_screen))
        file_menu.addAction(open_action)

        exit_action = QAction("Exit", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        about_action = QAction("About", self)
        about_action.triggered.connect(lambda: self.central_widget.setCurrentWidget(self.about_screen))
        help_menu.addAction(about_action)

    def create_footer(self):
        footer = QFrame()
        footer.setFrameShape(QFrame.Box)
        footer_layout = QHBoxLayout()
        footer_label = QLabel("© 2024 Elektriksel Test Donanımı Kontrol Sistemi. Tüm hakları saklıdır.")
        footer_layout.addWidget(footer_label, alignment=Qt.AlignCenter)
        footer.setLayout(footer_layout)
        footer.setFixedHeight(50)
        
        container_layout = QVBoxLayout()
        container_layout.addWidget(self.central_widget)
        container_layout.addWidget(footer)
        
        container_widget = QWidget()
        container_widget.setLayout(container_layout)
        
        self.setCentralWidget(container_widget)

    def create_main_screen(self):
        screen = QWidget()
        layout = QVBoxLayout()

        label = QLabel("Ana Ekran")
        layout.addWidget(label)
        layout.setAlignment(label, Qt.AlignCenter)

        btn_device_settings = QPushButton("Cihaz Ayarları")
        btn_device_settings.clicked.connect(lambda: self.central_widget.setCurrentWidget(self.device_settings_screen))
        layout.addWidget(btn_device_settings)
        layout.setAlignment(btn_device_settings, Qt.AlignCenter)

        btn_logging_settings = QPushButton("Veri Loglama Ayarları")
        btn_logging_settings.clicked.connect(lambda: self.central_widget.setCurrentWidget(self.logging_settings_screen))
        layout.addWidget(btn_logging_settings)
        layout.setAlignment(btn_logging_settings, Qt.AlignCenter)

        btn_manual_command = QPushButton("Manuel Komut Gönderme")
        btn_manual_command.clicked.connect(lambda: self.central_widget.setCurrentWidget(self.manual_command_screen))
        layout.addWidget(btn_manual_command)
        layout.setAlignment(btn_manual_command, Qt.AlignCenter)

        screen.setLayout(layout)
        return screen

    def create_device_settings_screen(self):
        screen = QWidget()
        layout = QVBoxLayout()

        label = QLabel("Cihaz Ayarları")
        layout.addWidget(label)
        layout.setAlignment(label, Qt.AlignCenter)

        self.device_file_input = QLineEdit()
        self.device_file_input.setPlaceholderText("XML Dosya Yolu")
        layout.addWidget(self.device_file_input)
        layout.setAlignment(self.device_file_input, Qt.AlignCenter)

        btn_browse = QPushButton("Dosya Seç")
        btn_browse.clicked.connect(self.browse_file)
        layout.addWidget(btn_browse)
        layout.setAlignment(btn_browse, Qt.AlignCenter)

        btn_load_xml = QPushButton("XML Yükle")
        btn_load_xml.clicked.connect(lambda: self.load_xml())
        layout.addWidget(btn_load_xml)
        layout.setAlignment(btn_load_xml, Qt.AlignCenter)

        btn_initialize = QPushButton("Initialize")
        btn_initialize.clicked.connect(lambda: self.start_initialize())
        layout.addWidget(btn_initialize)
        layout.setAlignment(btn_initialize, Qt.AlignCenter)

        self.device_status_list = QListWidget()
        layout.addWidget(self.device_status_list)
        layout.setAlignment(self.device_status_list, Qt.AlignCenter)

        btn_back = QPushButton("Geri")
        btn_back.clicked.connect(lambda: self.central_widget.setCurrentWidget(self.main_screen))
        layout.addWidget(btn_back)
        layout.setAlignment(btn_back, Qt.AlignCenter)

        self.progress_bar = QProgressBar()
        layout.addWidget(self.progress_bar)
        layout.setAlignment(self.progress_bar, Qt.AlignCenter)

        screen.setLayout(layout)
        return screen

    def create_logging_settings_screen(self):
        screen = QWidget()
        layout = QVBoxLayout()

        label = QLabel("Veri Loglama Ayarları")
        layout.addWidget(label)
        layout.setAlignment(label, Qt.AlignCenter)

        self.logging_interval_input = QLineEdit()
        self.logging_interval_input.setPlaceholderText("Loglama Aralığı (saniye)")
        layout.addWidget(self.logging_interval_input)
        layout.setAlignment(self.logging_interval_input, Qt.AlignCenter)

        btn_start_logging = QPushButton("Loglamayı Başlat")
        btn_start_logging.clicked.connect(lambda: self.start_logging())
        layout.addWidget(btn_start_logging)
        layout.setAlignment(btn_start_logging, Qt.AlignCenter)

        btn_stop_logging = QPushButton("Loglamayı Durdur")
        btn_stop_logging.clicked.connect(lambda: self.stop_logging())
        layout.addWidget(btn_stop_logging)
        layout.setAlignment(btn_stop_logging, Qt.AlignCenter)

        btn_back = QPushButton("Geri")
        btn_back.clicked.connect(lambda: self.central_widget.setCurrentWidget(self.main_screen))
        layout.addWidget(btn_back)
        layout.setAlignment(btn_back, Qt.AlignCenter)

        self.grid_layout = QGridLayout()
        layout.addLayout(self.grid_layout)

        screen.setLayout(layout)
        return screen

    def create_manual_command_screen(self):
        screen = QWidget()
        layout = QVBoxLayout()

        label = QLabel("Manuel Komut Gönderme")
        layout.addWidget(label)
        layout.setAlignment(label, Qt.AlignCenter)

        self.manual_ip_input = QLineEdit()
        self.manual_ip_input.setPlaceholderText("IP Adresi")
        layout.addWidget(self.manual_ip_input)
        layout.setAlignment(self.manual_ip_input, Qt.AlignCenter)

        self.manual_port_input = QLineEdit()
        self.manual_port_input.setPlaceholderText("Port")
        layout.addWidget(self.manual_port_input)
        layout.setAlignment(self.manual_port_input, Qt.AlignCenter)

        self.command_input = QLineEdit()
        self.command_input.setPlaceholderText("SCPI Komutu")
        layout.addWidget(self.command_input)
        layout.setAlignment(self.command_input, Qt.AlignCenter)

        self.command_list = QListWidget()
        self.load_commands()
        layout.addWidget(self.command_list)
        layout.setAlignment(self.command_list, Qt.AlignCenter)

        self.command_list.itemClicked.connect(self.on_command_selected)

        btn_connect_manual = QPushButton("Bağlan ve Komut Gönder")
        btn_connect_manual.clicked.connect(lambda: self.connect_and_send_manual_command())
        layout.addWidget(btn_connect_manual)
        layout.setAlignment(btn_connect_manual, Qt.AlignCenter)

        self.response_output = QTextEdit()
        self.response_output.setReadOnly(True)
        layout.addWidget(self.response_output)
        layout.setAlignment(self.response_output, Qt.AlignCenter)

        self.manual_command_status = QLabel()
        layout.addWidget(self.manual_command_status)
        layout.setAlignment(self.manual_command_status, Qt.AlignCenter)

        btn_back = QPushButton("Geri")
        btn_back.clicked.connect(lambda: self.central_widget.setCurrentWidget(self.main_screen))
        layout.addWidget(btn_back)
        layout.setAlignment(btn_back, Qt.AlignCenter)

        screen.setLayout(layout)
        return screen

    def create_about_screen(self):
        screen = QWidget()
        layout = QVBoxLayout()

        label = QLabel("Hakkında")
        layout.addWidget(label)
        layout.setAlignment(label, Qt.AlignCenter)

        about_text = QLabel("Bu uygulama, elektriksel test donanımlarını kontrol etmek ve loglamak için geliştirilmiştir.")
        layout.addWidget(about_text)
        layout.setAlignment(about_text, Qt.AlignCenter)

        btn_back = QPushButton("Geri")
        btn_back.clicked.connect(lambda: self.central_widget.setCurrentWidget(self.main_screen))
        layout.addWidget(btn_back)
        layout.setAlignment(btn_back, Qt.AlignCenter)

        screen.setLayout(layout)
        return screen

    def load_commands(self):
        commands = read_commands_from_document('commands.xml')  # Komutları dokümandan okuma
        for command in commands:
            self.command_list.addItem(command)

    def on_command_selected(self, item):
        self.command_input.setText(item.text())

    def browse_file(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(self, "XML Dosyası Seç", "", "XML Files (*.xml);;All Files (*)", options=options)
        if file_path:
            self.device_file_input.setText(file_path)

    def load_xml(self):
        file_path = self.device_file_input.text()
        if not file_path:
            QMessageBox.critical(self, "Hata", "Lütfen geçerli bir XML dosya yolu girin.")
            return

        try:
            self.devices = read_devices_from_xml(file_path)
            self.setup_logging_display(self.devices)  # XML yüklendikten sonra log ekranını güncelle
            QMessageBox.information(self, "Başarılı", "XML dosyası başarıyla yüklendi.")
        except FileNotFoundError:
            QMessageBox.critical(self, "Hata", f"XML dosyası bulunamadı: {file_path}")
        except ET.ParseError:
            QMessageBox.critical(self, "Hata", f"XML dosyası okunamadı: {file_path}")

    def start_initialize(self):
        if not self.devices:
            QMessageBox.critical(self, "Hata", "Lütfen önce XML dosyasını yükleyin.")
            return

        self.progress_bar.setValue(0)
        self.device_status_list.clear()

        self.worker = AsyncWorker(self.devices)
        self.worker.progress.connect(self.update_progress)
        self.worker.result.connect(self.show_result)
        self.worker.error.connect(self.show_error)
        self.worker.done.connect(self.initialization_done)

        self.loop = asyncio.get_event_loop()
        asyncio.ensure_future(self.worker.run())

    def update_progress(self, value):
        self.progress_bar.setValue(value)

    def show_result(self, message):
        self.device_status_list.addItem(message)

    def show_error(self, message):
        self.device_status_list.addItem(f"Hata: {message}")

    def initialization_done(self):
        QMessageBox.information(self, "Başarılı", "Initialize işlemi tamamlandı.")
        self.scpi_clients = self.worker.scpi_clients

    def setup_logging_display(self, devices):
        self.grid_layout = QGridLayout()
        row = 0
        col = 0

        for device in devices:
            if 'channels' in device:
                for channel in device.get('channels', []):
                    frame = QFrame()
                    frame.setFrameShape(QFrame.Box)
                    frame.setFixedSize(200, 50)
                    layout = QVBoxLayout()
                    label = QLabel(f"{channel['name']}")
                    value_label = QLabel("N/A")
                    layout.addWidget(label)
                    layout.addWidget(value_label)
                    frame.setLayout(layout)
                    self.grid_layout.addWidget(frame, row, col)
                    col += 1
                    if col == 8:
                        col = 0
                        row += 1

            if 'voltage_channels' in device:
                for channel in device.get('voltage_channels', []):
                    frame = QFrame()
                    frame.setFrameShape(QFrame.Box)
                    frame.setFixedSize(200, 50)
                    layout = QVBoxLayout()
                    label = QLabel(f"Voltage Channel {channel['number']}")
                    value_label = QLabel("N/A")
                    layout.addWidget(label)
                    layout.addWidget(value_label)
                    frame.setLayout(layout)
                    self.grid_layout.addWidget(frame, row, col)
                    col += 1
                    if col == 8:
                        col = 0
                        row += 1

            if 'current_channels' in device:
                for channel in device.get('current_channels', []):
                    frame = QFrame()
                    frame.setFrameShape(QFrame.Box)
                    frame.setFixedSize(200, 50)
                    layout = QVBoxLayout()
                    label = QLabel(f"Current Channel {channel['number']}")
                    value_label = QLabel("N/A")
                    layout.addWidget(label)
                    layout.addWidget(value_label)
                    frame.setLayout(layout)
                    self.grid_layout.addWidget(frame, row, col)
                    col += 1
                    if col == 8:
                        col = 0
                        row += 1

        # Log ekranındaki layout'u temizle ve yeni layout'u ekle
        self.central_widget.widget(2).layout().addLayout(self.grid_layout)

    def start_logging(self):
        asyncio.ensure_future(self.async_start_logging())

    async def async_start_logging(self):
        interval = int(self.logging_interval_input.text())
        print(f"Start logging every {interval} seconds")
        self.logging_timer = QTimer(self)
        self.logging_timer.timeout.connect(lambda: asyncio.ensure_future(self.log_data()))
        self.logging_timer.start(interval * 1000)

        with open(self.csv_file, 'w', newline='') as csvfile:
            fieldnames = ['Time', 'IP', 'Channel', 'Current', 'Voltage', 'Resistance']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()

    def stop_logging(self):
        if self.logging_timer:
            self.logging_timer.stop()
        print("Stop logging")

    async def log_data(self):
        with open(self.csv_file, 'a', newline='') as csvfile:
            writer = csv.writer(csvfile)
            for client in self.scpi_clients:
                try:
                    for channel in client.channels:
                        await client.send_command(f'CHAN {channel["number"]}')
                        current = await client.send_command('MEASure:CURRent?')
                        voltage = await client.send_command('MEASure:VOLTage?')
                        resistance = await client.send_command('MEASure:RESistance?')
                        print(f"{client.ip} - Channel {channel['number']} - Current: {current}, Voltage: {voltage}, Resistance: {resistance}")
                        writer.writerow([client.ip, channel['number'], current, voltage, resistance])
                        # Verileri ekranda güncelle
                        self.update_logging_display(client.ip, channel['number'], current, voltage, resistance)
                except Exception as e:
                    print(f"Error during logging from {client.ip}: {e}")

    def update_logging_display(self, ip, channel, current, voltage, resistance):
        for i in range(self.grid_layout.count()):
            frame = self.grid_layout.itemAt(i).widget()
            label = frame.findChild(QLabel, "")
            if label and label.text() == f"Channel {channel}":
                value_label = frame.findChild(QLabel, "", 1)
                if value_label:
                    value_label.setText(f"Current: {current}, Voltage: {voltage}, Resistance: {resistance}")

    def connect_and_send_manual_command(self):
        asyncio.ensure_future(self.async_connect_and_send_manual_command())

    async def async_connect_and_send_manual_command(self):
        ip = self.manual_ip_input.text()
        port = int(self.manual_port_input.text())
        command = self.command_input.text()

        if not ip or not port or not command:
            self.manual_command_status.setText("Lütfen geçerli bir IP, Port ve Komut girin.")
            return

        client = SCPIClient(ip, port)
        await client.connect()
        response = await client.send_command(command)
        if response:
            self.response_output.append(f"Komut: {command}\nYanıt: {response}\n")
            self.manual_command_status.setText(f"Komut gönderildi: {response}")
        else:
            self.response_output.append(f"Komut: {command}\nYanıt: (Yok)\n")
            self.manual_command_status.setText("Komut gönderilemedi.")
        client.close()

    def closeEvent(self, event):
        for client in self.scpi_clients:
            client.close()
        event.accept()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    loop = qasync.QEventLoop(app)
    asyncio.set_event_loop(loop)
    main_window = MainWindow()
    main_window.show()
    with loop:
        loop.run_forever()
