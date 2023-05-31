from main_window import Ui_MainWindow
from datetime import datetime

from PyQt5 import QtGui, QtWidgets
from PyQt5.QtWidgets import QMessageBox, QFileDialog
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPixmap, QCursor, QMouseEvent

import xlsxwriter
import threading
import requests
import sys
import os
import re


class Services:

    SERVICE_MAPPING = {
        "https://www.xoom.com/": "self.xoom_api",
        "https://www.worldremit.com/": "self.worldremit_api",
        "https://www.instarem.com/": "self.instarem_api",
        "https://app.currencyfair.com/": "self.currencyfair_api",
        "https://my.transfergo.com/": "self.transfergo_api",
        "https://secure.xendpay.com/": "self.xendpay_api",
        "https://wise.com/": "self.wise_api"
    }

    @staticmethod
    def instarem_api(source_country_code, source_currency_code, destination_country_code, destination_currency_code):
        base_url = 'https://www.instarem.com/api/v1/public/transaction/computed-value?'
        send_amount = 'source_amount=1000'

        url = base_url + '&'.join((
            'source_currency=' + source_currency_code,
            'destination_currency=' + destination_currency_code,
            'country_code=' + source_country_code,
            send_amount
        ))

        try:
            response = requests.get(url, timeout=5)
            data = response.json()
            return float(data['data']['fx_rate'])
        except Exception as e:
            print(f"Warning: API-instarem: {e}")
            return 0

    @staticmethod
    def currencyfair_api(source_country_code, source_currency_code, destination_country_code, destination_currency_code):
        base_url = 'https://app.currencyfair.com/calculator/quicktrade-quote?'
        send_amount = 'amount=40000'
        mode = 'mode=SELL'

        url = base_url + '&'.join((
            'depositCurrency=' + source_currency_code,
            'beneficiaryCurrency=' + destination_currency_code,
            send_amount,
            mode
        ))

        try:
            response = requests.get(url, timeout=5)
            data = response.json()
            return float(data['quote']['estimate']['rate'])
        except Exception as e:
            print(f"Warning: API-currencyfair: {e}")
            return 0

    @staticmethod
    def transfergo_api(source_country_code, source_currency_code, destination_country_code, destination_currency_code):
        base_url = 'https://my.transfergo.com/api/transfers/quote?'
        calculation_base = 'calculationBase=sendAmount'
        send_amount = 'amount=300.00'

        url = base_url + '&'.join((
            'fromCountryCode=' + source_country_code,
            'toCountryCode=' + destination_country_code,
            'fromCurrencyCode=' + source_currency_code,
            'toCurrencyCode=' + destination_currency_code,
            calculation_base,
            send_amount
        ))

        try:
            response = requests.get(url, timeout=5)
            data = response.json()
            return float(data['deliveryOptions']['standard']['paymentOptions']['bank']['quote']['rate'])
        except Exception as e:
            print(f"Warning: API-transfergo: {e}")
            return 0

    @staticmethod
    def xendpay_api(source_country_code, source_currency_code, destination_country_code, destination_currency_code):
        url = f'https://secure.xendpay.com/rate/XP/{source_currency_code}/{destination_currency_code}?cdc=true'
        try:
            response = requests.get(url, timeout=5)
            data = response.text
            return float(data)
        except Exception as e:
            print(f"Warning: API-xendpay: {e}")
            return 0

    @staticmethod
    def wise_api(source_country_code, source_currency_code, destination_country_code, destination_currency_code):
        base_url = 'https://wise.com/rates/history?'
        other = 'length=1&unit=day&resolution=hourly'

        url = base_url + '&'.join((
            'source=' + source_currency_code,
            'target=' + destination_currency_code,
            other
        ))

        try:
            response = requests.get(url, timeout=5)
            data = response.json()
            return data[-1]['value']
        except Exception as e:
            print(f"Warning: API-wise: {e}")
            return 0


class Main(Ui_MainWindow, Services):

    CONFIG = "currencies.config"
    NEW_LINE_FLAG = False
    WORKBOOK = None
    WORKSHEET = None
    PERCENTAGE = None
    THREAD = None
    SCRAPING_REQUESTS = []

    def __init__(self, __window) -> None:
        super(Main, self).setupUi(__window)
        super(Main, self).retranslateUi(__window)

        # FLAGS
        self.flag_maximized_geometry_before = (0, 0, 0, 0)
        self.flag_maximized = False
        self.dragging = False
        self.offset = None

        # EVENTS
        self.navigationPanel.mousePressEvent = self.mouse_press_event
        self.navigationPanel.mouseMoveEvent = self.mouse_move_event
        self.navigationPanel.mouseDoubleClickEvent = self.maximize_application
        self.closeButton.mousePressEvent = self.close_application
        self.minimizeButton.mousePressEvent = self.minimize_application
        self.maximizeButton.mousePressEvent = self.maximize_application

        # ICONS
        self.app_icon = QPixmap(self.resource_path('icons/main.ico'))
        self.close_icon = QPixmap(self.resource_path('icons/close.png'))
        self.expand_icon = QPixmap(self.resource_path('icons/expand.png'))
        self.minus_icon = QPixmap(self.resource_path('icons/minus.png'))

        # SETTING ICONS
        self.closeButton.setPixmap(self.close_icon)
        self.closeButton.setCursor(QCursor(Qt.PointingHandCursor))
        self.maximizeButton.setPixmap(self.expand_icon)
        self.maximizeButton.setCursor(QCursor(Qt.PointingHandCursor))
        self.minimizeButton.setPixmap(self.minus_icon)
        self.minimizeButton.setCursor(QCursor(Qt.PointingHandCursor))
        self.appLogo.setPixmap(self.app_icon)

        # SETTING CONNECTIONS
        self.set_connections()
        self.show_welcome_message()

    def set_connections(self) -> None:
        self.browseButton.clicked.connect(self.browse_folders)
        self.generateButton.clicked.connect(self.run_generate_report)
        self.suggestedCurrenciesTable.itemDoubleClicked.connect(self.add_currency)
        self.reportFileDirectoryField.setText(os.getcwd())
        self.add_currencies_from_config()

    def add_currencies_from_config(self) -> None:
        index = 0
        config_path = self.resource_path(self.CONFIG)
        with open(config_path, "r") as config:
            for line in config:
                self.suggestedCurrenciesTable.insertRow(index)
                self.suggestedCurrenciesTable.setItem(index, 0, QtWidgets.QTableWidgetItem(line.strip()))
                index += 1

    def add_currency(self) -> None:
        currency_row = self.suggestedCurrenciesTable.currentRow()
        currency = self.suggestedCurrenciesTable.item(currency_row, 0).text().split(' - ')[0]

        if self.NEW_LINE_FLAG:
            current_text = self.currencyAndCountriesField.toPlainText()
            self.currencyAndCountriesField.setText(current_text + currency)
            self.NEW_LINE_FLAG = False
        else:
            self.currencyAndCountriesField.append(currency + '\t')
            self.NEW_LINE_FLAG = True

    def init_excel(self) -> None:
        location = self.reportFileDirectoryField.text()
        current_date = datetime.today().strftime('%Y_%m_%d__%H_%M_%S')

        self.WORKBOOK = xlsxwriter.Workbook(location + r'\report_' + current_date + '.xlsx')
        self.WORKSHEET = self.WORKBOOK.add_worksheet()

        self.WORKSHEET.set_row(0, 75)
        self.WORKSHEET.set_column(4, 100, 15)

        default_format = self.WORKBOOK.add_format({
            'text_wrap': True,
            'align': 'left',
            'valign': 'top',
        })

        link_format = self.WORKBOOK.add_format({
            'color': 'blue',
            'underline': True,
            'text_wrap': True,
            'align': 'left',
            'valign': 'top',
        })

        headers = [
            'Source Country Code',
            'Source Currency Code',
            'Destination Country Code',
            'Destination Currency Code',
        ]

        for link in range(1, 11):
            link_checkbox: QtWidgets.QCheckBox
            link_checkbox = eval(f"self.page_{link}")
            if link_checkbox.isChecked():
                self.SCRAPING_REQUESTS.append(link_checkbox.text())

        index = 0
        cell_format = default_format
        for header in headers + self.SCRAPING_REQUESTS:
            if index == 4:
                cell_format = link_format
            self.WORKSHEET.write(0, index, header, cell_format)
            index += 1

    def browse_folders(self) -> None:
        selected_directory = QFileDialog().getExistingDirectory(directory=os.getcwd())
        if selected_directory:
            self.reportFileDirectoryField.setText(selected_directory)

    def run_generate_report(self) -> None:
        text = self.currencyAndCountriesField.toPlainText()
        errors = self.check_errors_in_input_text(text)
        if errors:
            self.show_error_message()
            return

        t1 = threading.Thread(target=self.generate_report, args=(text,))
        t1.start()

    def generate_report(self, text: str) -> None:
        self.progressBar.setValue(0)
        self.generateButton.setEnabled(False)
        self.browseButton.setEnabled(False)

        all_currencies = text.split('\n')

        self.init_excel()
        self.PERCENTAGE = 100 // len(all_currencies) // len(self.SCRAPING_REQUESTS)

        for line_index, each_currency in enumerate(all_currencies):
            if not each_currency:
                continue

            m = re.match(r'(\w+)'  # source_country_code: m.group(1).upper()
                         r'\s+(\w+)'  # source_currency_code: m.group(2).upper()
                         r'\s+(\w+)'  # destination_country_code: m.group(3).upper()
                         r'\s+(\w+)'  # destination_currency_code = m.group(4).upper()
                         r'', each_currency, re.X)

            rates = [m.group(1).upper(), m.group(2).upper(), m.group(3).upper(), m.group(4).upper()]

            for page in self.SCRAPING_REQUESTS:
                scrapper = self.SERVICE_MAPPING[page]
                rates.append(eval(f"{scrapper}(rates[0], rates[1], rates[2], rates[3])"))
                self.update_progress()

            for column_index, each_rate in enumerate(rates):
                if column_index < 4:
                    self.WORKSHEET.write(line_index + 1, column_index, each_rate)
                else:
                    self.WORKSHEET.write(line_index + 1, column_index, float(format(float(each_rate), ".2f")))

        self.WORKBOOK.close()
        self.SCRAPING_REQUESTS.clear()
        self.generateButton.setEnabled(True)
        self.browseButton.setEnabled(True)
        self.progressBar.setValue(100)

    def update_progress(self) -> None:
        current_per = int(self.progressBar.value())
        final_per = current_per + self.PERCENTAGE + 1
        self.progressBar.setValue(final_per)

    def mouse_press_event(self, event: QMouseEvent) -> None:
        if event.button() == Qt.LeftButton:
            if event.pos() in self.navigationPanel.rect().translated(self.navigationPanel.pos()):
                self.dragging = True
                self.offset = event.pos()
                event.accept()

    def mouse_move_event(self, event: QMouseEvent) -> None:
        if self.dragging:
            MainWindow.move(event.globalPos() - self.offset)
            event.accept()

    def maximize_application(self, event: QMouseEvent) -> None:
        if self.flag_maximized:
            x, y, width, height = self.flag_maximized_geometry_before
            self.flag_maximized = False
            MainWindow.showNormal()
            MainWindow.setGeometry(x, y, width, height)
        else:
            self.flag_maximized_geometry_before = MainWindow.geometry().getRect()
            self.flag_maximized = True
            MainWindow.showMaximized()
        event.accept()

    @staticmethod
    def check_errors_in_input_text(text: str) -> bool:
        pattern = r"(\w+)\s+(\w+)\s+(\w+)\s+(\w+)\s*"
        m = re.search(rf'^{pattern}(\n{pattern})*$', text)
        return not bool(m)

    @staticmethod
    def show_error_message() -> None:
        msg = QMessageBox()
        msg.setWindowTitle("Oops - Syntax Error")
        msg.setWindowIcon(QtGui.QIcon(Main.resource_path("./icons/main.ico")))
        msg.setIcon(QMessageBox.Critical)
        msg.setStyleSheet("border-radius: 10px;\n"
                          "background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:1, stop:"
                          "0 rgba(148, 202, 236, 255), stop:1 rgba(146, 153, 255, 255));\n"
                          "font: 10pt \"Gill Sans MT\";")
        msg.setText("Please use the following syntax:\n\n"
                    "(Country Code 1) (Currency 1) TAB (Country Code 1) (Currency 2)\t\n"
                    "(Country Code 2) (Currency 2) TAB (Country Code 1) (Currency 2)\t\n"
                    "...\n\n"
                    "The output file will contain rates for each of the currency lines.\n"
                    "Double click on the currency to add it."
                    )
        msg.exec_()

    @staticmethod
    def close_application(event: QMouseEvent) -> None:
        MainWindow.close()
        event.accept()

    @staticmethod
    def minimize_application(event: QMouseEvent) -> None:
        MainWindow.showMinimized()
        event.accept()

    @staticmethod
    def resource_path(relative_path: str) -> str:
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath("../../Desktop/F/v2.0")

        return os.path.join(base_path, relative_path)

    @staticmethod
    def show_welcome_message():
        msg = QMessageBox()
        msg.setWindowTitle("Welcome!")
        msg.setWindowIcon(QtGui.QIcon(Main.resource_path("./icons/main.ico")))
        msg.setIcon(QMessageBox.Information)
        msg.setStyleSheet("border-radius: 10px;\n"
                          "background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:1, stop:"
                          "0 rgba(148, 202, 236, 255), stop:1 rgba(146, 153, 255, 255));\n"
                          "font: 10pt \"Gill Sans MT\";")
        msg.setText("Welcome to my demo application!\n\n"
                    "Your involvement is incredibly valuable to me. As the sole developer, "
                    "your insights and perspectives will play a crucial role in shaping the future of this application."
                    " I'm eager to hear your thoughts, so please don't hesitate to reach out to me through the GitHub "
                    "repository or any other feedback channels.\n\n"
                    "GitHub: https://github.com/lgrigor/Rate-Scrapper\n\n"
                    "Thank you for joining me on this journey. \nTogether, let's make something extraordinary.\n\n"
                    "Kind regards, Levon")
        msg.exec_()


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Main(MainWindow)

    MainWindow.setWindowTitle("Rate Scrapper")
    MainWindow.setWindowFlags(Qt.FramelessWindowHint)
    MainWindow.show()

    sys.exit(app.exec_())
