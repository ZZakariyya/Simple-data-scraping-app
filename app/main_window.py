from PySide6.QtWidgets import QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget, QFileDialog, QLabel, QLineEdit, QTextEdit
from PySide6.QtCore import QThread, Signal
from PySide6.QtGui import QIcon
import sys
import pandas as pd

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager

class WorkerThread(QThread):
    update_signal = Signal(str)

    def __init__(self, excel_path, excel_save_path):
        super().__init__()
        self.excel_path = excel_path
        self.excel_save_path = excel_save_path
        # self.driver_path = driver_path

    def run(self):
        try:
            chrome_driver_path = ChromeDriverManager().install()

            service = Service(executable_path=chrome_driver_path)
            options = webdriver.ChromeOptions()
            options.add_argument("--incognito")
            options.add_argument("--start-maximized")
            driver = webdriver.Chrome(service=service, options=options)


            website_url = 'https://emas.sosial.gov.az/warParticipantVictory'
            driver.get(website_url)

            df = pd.read_excel(self.excel_path)
            for index, row in df.iterrows():
                name_to_check = row['FIN']
                input_field_xpath = '//*[@id="app"]/div/div[4]/div/div[1]/div/div[2]/div/div/div[1]/input'
                input_element = driver.find_element(By.XPATH, input_field_xpath)
                input_element.clear()
                input_element.send_keys(name_to_check)
                input_element.send_keys(Keys.ENTER)

                try:
                    status_element_xpath = '//*[@id="app"]/div/div[4]/div/div[1]/div/div[2]/div/div/div[3]/div/div/div'
                    element = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, status_element_xpath)))

                    status_text = element.text.split('Status:')[-1].strip()
                    df.at[index, 'Status'] = status_text
                    self.update_signal.emit(f"Processed {name_to_check}: {status_text}")

                except TimeoutException:
                    df.at[index, 'Status'] = "Not Found"
                    self.update_signal.emit(f"Processed {name_to_check}: Not Found")

            df.to_excel(self.excel_save_path, index=False)
            self.update_signal.emit("New Excel Save file is created!")

        except Exception as e:
            self.update_signal.emit(f"Error: {e}")

        finally:
            if driver:
                driver.quit()



class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Excel Processing with Selenium")
        self.setGeometry(100, 100, 800, 600)

        self.setStyleSheet("background-color: lightblue;")

        self.setWindowIcon(QIcon('E:/scraping_spider/spider1.jpg'))

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        self.excel_file_line_edit = QLineEdit()
        self.excel_file_line_edit.setPlaceholderText("Path to open Excel file...")
        layout.addWidget(self.excel_file_line_edit)

        self.csv_file_line_edit = QLineEdit()
        self.csv_file_line_edit.setPlaceholderText("Path to save Excel file...")
        layout.addWidget(self.csv_file_line_edit)

        # self.driver_path_line_edit = QLineEdit()
        # self.driver_path_line_edit.setPlaceholderText("Path to ChromeDriver...")
        # layout.addWidget(self.driver_path_line_edit)

        self.open_excel_button = QPushButton("Open Excel File")
        self.open_excel_button.clicked.connect(self.open_excel_dialog)
        layout.addWidget(self.open_excel_button)

        self.open_csv_button = QPushButton("Set Excel Save Path")
        self.open_csv_button.clicked.connect(self.open_csv_dialog)
        layout.addWidget(self.open_csv_button)

        # self.driver_path_button = QPushButton("Set Driver Path")
        # self.driver_path_button.clicked.connect(self.driver_path_dialog)
        # layout.addWidget(self.driver_path_button)

        self.start_button = QPushButton("Start Processing")
        self.start_button.clicked.connect(self.start_process)
        layout.addWidget(self.start_button)

        self.status_text_edit = QTextEdit()
        self.status_text_edit.setReadOnly(True)
        layout.addWidget(self.status_text_edit)

    def open_excel_dialog(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "Open Excel File", "", "Excel Files (*.xlsx *.xls)")
        if file_name:
            self.excel_file_line_edit.setText(file_name)

    def open_csv_dialog(self):
        file_name, _ = QFileDialog.getSaveFileName(self, "Save Excel File", "", "Excel Files (*.xlsx)")
        if file_name:
            self.csv_file_line_edit.setText(file_name)

    # def driver_path_dialog(self):
    #     file_name, _ = QFileDialog.getOpenFileName(self, "Save your data path for your chromedriver")
    #     if file_name:
    #         self.driver_path_line_edit.setText(file_name)

    def start_process(self):
        excel_path = self.excel_file_line_edit.text()
        excel_save_path = self.csv_file_line_edit.text()
        # driver_path = self.driver_path_line_edit.text()

        if not excel_path or not excel_save_path:
            self.status_text_edit.append("Please specify all file paths.")
            return

        self.thread = WorkerThread(excel_path, excel_save_path)
        self.thread.update_signal.connect(self.update_status)
        self.thread.start()

    def update_status(self, message):
        self.status_text_edit.append(message)

if __name__ == "__main__":
    app = QApplication([])
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


app = QApplication([])
window = MainWindow()
window.show()
sys.exit(app.exec())
