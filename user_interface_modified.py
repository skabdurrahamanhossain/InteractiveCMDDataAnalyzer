import os
import threading
from datetime import datetime, timedelta
import sys
import json
import subprocess
import re
import time
import pandas as pd
from pandas.errors import EmptyDataError
import signal

import numpy as np
from PyQt5.QtWidgets import QApplication, QMessageBox, QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout, \
    QHBoxLayout, QFileDialog


class TestSetupWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.sensor_part_number = None
        self.min_data_rate_limit = None
        self.max_data_rate_limit = None
        self.max_std_dev_limit = None
        self.test_run_time = None
        self.comserver_path = None
        self.test_result_path = None

        self.output_window = None

        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # Buttons for Import and Export
        import_button = QPushButton("Import", self)
        import_button.clicked.connect(self.import_configuration)
        export_button = QPushButton("Export", self)
        export_button.clicked.connect(self.export_configuration)

        button_layout = QHBoxLayout()
        button_layout.addWidget(import_button)
        button_layout.addWidget(export_button)
        layout.addLayout(button_layout)

        # Input fields
        layout.addLayout(self.create_input_layout("Sensor Part Number:", "sensor_part_number"))
        layout.addLayout(self.create_input_layout("Minimum Data Rate Limit :", "min_data_rate_limit", "MB/s"))
        layout.addLayout(self.create_input_layout("Maximum Data Rate Limit :", "max_data_rate_limit", "MB/s"))
        layout.addLayout(self.create_input_layout("Maximum Standard Deviation Limit:", "max_std_dev_limit"))
        layout.addLayout(self.create_input_layout("Test Run Time:", "test_run_time", "s"))
        layout.addLayout(self.create_input_browse_path_layout("COMSERVER Path:", "comserver_path"))
        layout.addLayout(self.create_output_browse_path_layout("Test Result Path:", "test_result_path"))

        # Start Button
        start_button = QPushButton("Next", self)
        start_button.clicked.connect(self.next_window)
        layout.addWidget(start_button)

        self.setLayout(layout)
        self.setWindowTitle("Test Parameter Setup Window")
        # Set fixed width and height for the window (adjust the values as needed)
        self.setFixedSize(700, 500)
        self.show()

    def create_input_layout(self, label_text, attribute_name, post_label_text=""):
        input_layout = QHBoxLayout()

        # Label
        label = QLabel(label_text)
        input_layout.addWidget(label)

        # Line Edit
        line_edit = QLineEdit(self)
        setattr(self, attribute_name + "_entry", line_edit)
        input_layout.addWidget(line_edit)
        post_label = QLabel(post_label_text)
        input_layout.addWidget(post_label)

        return input_layout

    def create_input_browse_path_layout(self, label_text, attribute_name):
        input_layout = QHBoxLayout()

        # Label
        label = QLabel(label_text)
        input_layout.addWidget(label)

        # Line Edit
        line_edit = QLineEdit(self)
        setattr(self, attribute_name + "_entry", line_edit)
        input_layout.addWidget(line_edit)
        comserver_button = QPushButton("...", self)
        comserver_button.clicked.connect(self.browse_comserver_path)
        input_layout.addWidget(comserver_button)

        return input_layout

    def create_output_browse_path_layout(self, label_text, attribute_name):
        input_layout = QHBoxLayout()

        # Label
        label = QLabel(label_text)
        input_layout.addWidget(label)

        # Line Edit
        line_edit = QLineEdit(self)
        setattr(self, attribute_name + "_entry", line_edit)
        input_layout.addWidget(line_edit)
        result_button = QPushButton("...", self)
        result_button.clicked.connect(self.browse_output_path)
        input_layout.addWidget(result_button)

        return input_layout

    def next_window(self):
        # Get values from input fields
        self.sensor_part_number = self.sensor_part_number_entry.text()
        self.min_data_rate_limit = float(self.min_data_rate_limit_entry.text())
        self.max_data_rate_limit = float(self.max_data_rate_limit_entry.text())
        self.max_std_dev_limit = float(self.max_std_dev_limit_entry.text())
        self.test_run_time = int(self.test_run_time_entry.text())
        self.comserver_path = self.comserver_path_entry.text()
        self.test_result_path = self.test_result_path_entry.text()
        if self.output_window:
            if self.output_window.comserver_process and self.output_window.comserver_process.poll() is None:
                self.output_window.stop_comserver()
            self.output_window = None
        # Open the output window
        self.output_window = TestOutputWindow(self)
        self.output_window.show()

        # Perform any other actions needed to start the test

    def refresh_setup_window(self):
        # Get values from input fields
        self.sensor_part_number = self.sensor_part_number_entry.text()
        self.min_data_rate_limit = float(self.min_data_rate_limit_entry.text())
        self.max_data_rate_limit = float(self.max_data_rate_limit_entry.text())
        # self.min_std_dev_limit = float(self.min_std_dev_limit_entry.text())
        self.max_std_dev_limit = float(self.max_std_dev_limit_entry.text())
        self.test_run_time = int(self.test_run_time_entry.text())
        self.comserver_path = self.comserver_path_entry.text()
        self.test_result_path = self.test_result_path_entry.text()

    def import_configuration(self):
        file_dialog = QFileDialog(self)
        file_path, _ = file_dialog.getOpenFileName(self, "Import Configuration", "", "JSON Files (*.json)")
        if file_path:
            with open(file_path, "r") as file:
                data = json.load(file)
                # Set values to input fields
                self.sensor_part_number_entry.setText(data.get("sensor_part_number", ""))
                self.min_data_rate_limit_entry.setText(str(data.get("min_data_rate_limit", "")))
                self.max_data_rate_limit_entry.setText(str(data.get("max_data_rate_limit", "")))
                self.max_std_dev_limit_entry.setText(str(data.get("max_std_dev_limit", "")))
                self.test_run_time_entry.setText(str(data.get("test_run_time", "")))
                self.comserver_path_entry.setText(str(data.get("comserver_path", "")))
                self.test_result_path_entry.setText(str(data.get("test_result_path", "")))

    def export_configuration(self):
        file_dialog = QFileDialog(self)
        file_path, _ = file_dialog.getSaveFileName(self, "Export Configuration", "", "JSON Files (*.json)")
        if file_path:
            data = {
                "sensor_part_number": self.sensor_part_number_entry.text(),
                "min_data_rate_limit": float(self.min_data_rate_limit_entry.text()),
                "max_data_rate_limit": float(self.max_data_rate_limit_entry.text()),
                "max_std_dev_limit": float(self.max_std_dev_limit_entry.text()),
                "test_run_time": int(self.test_run_time_entry.text()),
                "comserver_path": self.comserver_path_entry.text(),
                "test_result_path": self.test_result_path_entry.text()
            }
            with open(file_path, "w") as file:
                json.dump(data, file)

    def browse_comserver_path(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        comserver_path, _ = QFileDialog.getOpenFileName(self, "Select COMSERVER Path", "",
                                                        "batch Files (*.bat)", options=options)

        if comserver_path:
            self.comserver_path_entry.setText(comserver_path)

    def browse_output_path(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        output_path, _ = QFileDialog.getOpenFileName(self, "Select output Path", "",
                                                     "csv Files (*.csv)", options=options)

        if output_path:
            self.test_result_path_entry.setText(output_path)


class TestOutputWindow(QWidget):
    def __init__(self, setup_window):
        super().__init__()

        self.setup_window = setup_window

        # self.test_status_label = None
        self.sensor_part_number = None
        self.a2c_number_input = None
        self.test_output_label = None
        self.average_data_rate_label = None
        self.std_deviation_label = None
        self.errors_label = None
        self.test_run_time_label = None
        self.comserver_status_label = None

        self.start_comserver_button = None
        self.stop_comserver_button = None

        self.comserver_process = None  # To store the process
        self.comserver_running = False

        self.init_ui()

    def change_button_color(self):
        if not self.comserver_running:
            QMessageBox.warning(self, "COMSERVER Not Running", "Please start COMSERVER before starting the test.")
            return
        # Check if A2C number is filled
        a2c_number = self.a2c_number_input.text().strip()
        if not a2c_number:
            QMessageBox.warning(self, "Sensor Serial Number Not Filled",
                                "Please enter the Sensor Serial Number before starting the test.")
            return
        self.start_button.setStyleSheet("background-color: green;")
        self.start_button.setText("Running...")
        self.setup_window.refresh_setup_window()
        self.refresh_data()

    def restore_button_color(self):
        self.start_button.setStyleSheet("")
        self.start_button.setText("Start Test")

    def init_ui(self):
        layout = QVBoxLayout()

        # Buttons to Start and Stop COMSERVER
        self.start_comserver_button = QPushButton("Start Comserver", self)
        self.start_comserver_button.clicked.connect(self.start_comserver)
        self.stop_comserver_button = QPushButton("Stop Comserver", self)
        self.stop_comserver_button.clicked.connect(self.stop_comserver)
        self.stop_comserver_button.pressed.connect(self.stop_comserver_pressed)

        button_layout = QHBoxLayout()
        button_layout.addWidget(self.start_comserver_button)
        button_layout.addWidget(self.stop_comserver_button)
        layout.addLayout(button_layout)

        self.sensor_part_number = QLabel(f"Sensor Part Number: {self.setup_window.sensor_part_number}", self)
        layout.addWidget(self.sensor_part_number)

        # Input field for A2C number
        self.a2c_number_input = QLineEdit(self)
        self.a2c_number_input.setPlaceholderText("Sensor Serial Number")  # Placeholder text
        layout.addWidget(self.a2c_number_input)

        self.test_output_label = QLabel("Test Output: No Results", self)
        layout.addWidget(self.test_output_label)

        self.average_data_rate_label = QLabel("Average Data Rate: 0 MB/s", self)
        layout.addWidget(self.average_data_rate_label)

        self.std_deviation_label = QLabel("Standard Deviation: 0", self)
        layout.addWidget(self.std_deviation_label)

        self.errors_label = QLabel("Errors: 0", self)
        layout.addWidget(self.errors_label)

        self.is_button_clicked = False
        # Buttons to Start test
        self.start_button = QPushButton("Start Test", self)
        self.start_button.clicked.connect(self.start_test)
        self.start_button.pressed.connect(self.change_button_color)
        self.start_button.released.connect(self.restore_button_color)

        start_button_layout = QHBoxLayout()
        start_button_layout.addWidget(self.start_button)
        layout.addLayout(start_button_layout)

        self.test_run_time_label = QLabel(f"Test Run Time: {self.setup_window.test_run_time} s", self)

        self.comserver_status_label = QLabel("COMSERVER Status: Not Running", self)
        # layout.addWidget(self.comserver_status_label)
        output_labels_layout = QHBoxLayout()  # New horizontal layout for these labels

        # Add comserver_status_label and test_run_time_label to the new horizontal layout
        output_labels_layout.addWidget(self.comserver_status_label)
        output_labels_layout.addWidget(self.test_run_time_label)

        # Add the new horizontal layout to the main layout
        layout.addLayout(output_labels_layout)

        self.setLayout(layout)
        self.setWindowTitle("Test Output Window")

        # Set fixed width and height for the window (adjust the values as needed)
        self.setFixedSize(600, 400)

    def extract_value(self, line):
        # Define a regular expression pattern to match the desired value
        pattern = r'rx=(\d+\.\d+) MB/s'

        # Use re.search to find the pattern in the line
        match = re.search(pattern, line)

        # If a match is found, extract the value
        if match:
            return match.group(1)
        else:
            return None

    def is_csv_file_empty(self, file_path):
        try:
            df = pd.read_csv(file_path)
            if len(df) == 0:
                return True
            else:
                return False
        except EmptyDataError:
            return True

    def update_excel(self, excel_file, Test_starting_time, Sensor_name, A2C_number, Minimum_Data_Rate_Limit,
                     Maximum_Data_Rate_Limit, Average_data_rate,
                     Maximum_Standard_Deviation_Limit, Standard_Deviation, Errors, Test_status,
                     Test_end_time):
        is_empty = self.is_csv_file_empty(excel_file)
        # Load the Excel file into a Pandas DataFrame
        if is_empty:
            # creating empty dataframe
            df_updated = pd.DataFrame({'Test starting time': [Test_starting_time], 'Sensor name': [''],
                                       'Sensor Serial Number': [''], 'Minimum Data Rate Limit': [''],
                                       'Maximum Data Rate Limit': [''], 'Average data rate': [''],
                                       'Maximum Standard Deviation Limit': [''], 'Standard Deviation': [''],
                                       'Errors': [''], 'Test status': [''], 'Test end time': [Test_end_time]})
        else:
            df = pd.read_csv(excel_file)

        # Create a new DataFrame with the new data to be inserted
        new_data = {'Test starting time': [Test_starting_time],
                    'Sensor name': [Sensor_name],
                    'Sensor Serial Number': [A2C_number],
                    'Minimum Data Rate Limit': [Minimum_Data_Rate_Limit],
                    'Maximum Data Rate Limit': [Maximum_Data_Rate_Limit],
                    'Average data rate': [Average_data_rate],
                    'Maximum Standard Deviation Limit': [Maximum_Standard_Deviation_Limit],
                    'Standard Deviation': [Standard_Deviation],
                    'Errors': [Errors],
                    'Test status': [Test_status],
                    'Test end time': [Test_end_time]
                    }
        df_new = pd.DataFrame(new_data)
        if is_empty:
            # Append the new data to the empty DataFrame
            df_updated.update(df_new)
        else:
            # Append the new data to the existing DataFrame
            df_updated = pd.concat([df, df_new], ignore_index=True)

        # Save the updated DataFrame to the Excel file
        df_updated.to_csv(excel_file, index=False)

    def refresh_data(self):
        # Update labels with the latest data from setup window
        self.sensor_part_number.setText(f"Sensor Part Number: {self.setup_window.sensor_part_number}")
        self.test_run_time_label.setText(f"Test Run Time: {self.setup_window.test_run_time} s")

        # Clear all output labels
        self.average_data_rate_label.setText(f"Average Data Rate: 0 MB/s")
        self.std_deviation_label.setText(f"Standard Deviation: 0")
        self.errors_label.setText(f"Errors: 0")
        self.test_output_label.setText(f"Test Output: Running...")


    def start_test(self):
        # if not self.comserver_running:
        #     QMessageBox.warning(self, "COMSERVER Not Running", "Please start COMSERVER before starting the test.")
        #     return
        # # Check if A2C number is filled
        # a2c_number = self.a2c_number_input.text().strip()
        # if not a2c_number:
        #     QMessageBox.warning(self, "Sensor Serial Number Not Filled", "Please enter the Sensor Serial Number before starting the test.")
        #     return
        self.is_button_clicked = True
        # Refresh data from the setup window
        # self.setup_window.refresh_setup_window()
        # # self.refresh_data()

        startTime = datetime.now()
        currTime = time.time()
        test_run_time_label = self.setup_window.test_run_time
        self.test_run_time_label.setText(f"Test Run Time: {self.setup_window.test_run_time} s")
        endTime = startTime + timedelta(seconds=self.setup_window.test_run_time)
        errors_warnings = 0
        iteration = 0
        sum_of_data_rate = 0.0
        rx_value_list = []
        average_data_rate = 0
        standard_deviation = 0
        Test_status = ''
        previous_line = ''
        same_line_cnt = 0
        while datetime.now() < endTime and self.comserver_process:
            fd = self.comserver_process.stdout.fileno()
            os.set_blocking(fd, False)
            fe = self.comserver_process.stderr.fileno()
            os.set_blocking(fe, False)
            output_line = self.comserver_process.stderr.readline()
            if output_line == '':
                # time.sleep(4)
                counter = 0
                while output_line == '' and datetime.now() < endTime:
                    output_line = self.comserver_process.stdout.readline()
                    if output_line == '':
                        counter += 1
                        output_line = self.comserver_process.stderr.readline()
                        # if counter>300000:
                        #     errors_warnings += 1
                        #     break
                    else:
                        break

            if previous_line != output_line and len(output_line) > 3:
                previous_line = output_line
                test_run_time_label -= 1  # Decrease the variable
                # currTime = time.time()  # Reset the start time
                # self.test_run_time_label.setText(f"Test Run Time: {test_run_time_label} s")
                if output_line == '' and self.comserver_process.poll() is not None:
                    errors_warnings += 1
                    # break
                if output_line:
                    rx_value = self.extract_value(output_line)
                    if rx_value is not None:
                        iteration += 1
                        sum_of_data_rate += float(rx_value)
                        rx_value_list.append(float(rx_value))
                    else:
                        errors_warnings += 1
                        # break
            else:
                same_line_cnt += 1
                if same_line_cnt > 1:
                    errors_warnings += 1
                    # break

        if iteration > 0:
            average_data_rate = sum_of_data_rate / iteration
            standard_deviation = np.std(rx_value_list)
        if average_data_rate < self.setup_window.min_data_rate_limit or average_data_rate > self.setup_window.max_data_rate_limit:
            self.average_data_rate_label.setText(
                f"Average Data Rate: <span style='color: red;'>{round(average_data_rate, 2)} MB/s</span>")
        else:
            self.average_data_rate_label.setText(f"Average Data Rate: {round(average_data_rate, 2)} MB/s")
        if standard_deviation > self.setup_window.max_std_dev_limit:
            self.std_deviation_label.setText(
                f"Standard Deviation: <span style='color: red;'>{round(float(standard_deviation), 2)}</span>")
        else:
            self.std_deviation_label.setText(f"Standard Deviation: {round(float(standard_deviation), 2)}")
        if errors_warnings > 0:
            self.errors_label.setText(f"Errors: <span style='color: red;'>{errors_warnings}</span>")
        else:
            self.errors_label.setText(f"Errors: {errors_warnings}")
        if errors_warnings > 0 or average_data_rate < self.setup_window.min_data_rate_limit or average_data_rate > self.setup_window.max_data_rate_limit or standard_deviation > self.setup_window.max_std_dev_limit:
            self.test_output_label.setText(f"<div style='padding: 5px;'>Test Output: <span style='background-color: "
                                           f"red; color: black; font-weight: bold;'>FAIL</span></div>")
            result_msg = f"<div style='padding: 5px;'>Test Output: <span style='background-color: red; color: black; font-weight: bold;'>FAIL</span></div>"
            Test_status = 'FAIL'
        else:
            self.test_output_label.setText(f"<div style='padding: 5px;'>Test Output: <span style='background-color: "
                                           f"green; color: black; font-weight: bold;'>PASS</span></div>")
            result_msg = f"<div style='padding: 5px;'>Test Output: <span style='background-color: green; color: black; font-weight: bold;'>PASS</span></div>"
            Test_status = 'PASS'
        result_msg += f"Average Data Rate: {round(average_data_rate, 2)} MB/s<br>"
        result_msg += f"Standard Deviation: {round(float(standard_deviation), 2)}<br>"
        result_msg += f"Errors: {errors_warnings}\n\n"

        self.update_excel(self.setup_window.test_result_path, startTime, self.setup_window.sensor_part_number,
                          self.a2c_number_input.text(),
                          self.setup_window.min_data_rate_limit,
                          self.setup_window.max_data_rate_limit, round(average_data_rate, 2),
                          self.setup_window.max_std_dev_limit, round(float(standard_deviation), 2), errors_warnings,
                          Test_status,
                          endTime)
        self.is_button_clicked = False
        self.a2c_number_input.clear()
        self.start_thread_to_ignore_unnecessary_lines()

    def run_loop(self):
        counter = 0
        fd = self.comserver_process.stdout.fileno()
        os.set_blocking(fd, False)
        fe = self.comserver_process.stderr.fileno()
        os.set_blocking(fe, False)
        while not self.is_button_clicked and self.comserver_running and self.comserver_process:
            current_line = self.comserver_process.stdout.readline()
            output_line = self.comserver_process.stderr.readline()
            if len(current_line) < 42:
                counter += 1
            else:
                counter = 0
            time.sleep(0.1)
            if counter > 20:
                # self.comserver_status_label.setText("COMSERVER Status: Sensor not connected")
                self.comserver_status_label.setText(
                    "COMSERVER Status: <span style='color: #9C5700;'>Sensor not connected</span>")
            else:
                # self.comserver_status_label.setText("COMSERVER Status: Running")
                self.comserver_status_label.setText("COMSERVER Status: <span style='color: green;'>Running</span>")

    def start_thread_to_ignore_unnecessary_lines(self):
        loop_thread = threading.Thread(target=self.run_loop)
        loop_thread.start()

    def start_comserver(self):
        try:
            # Check if COMSERVER is already running
            if self.comserver_process and self.comserver_process.poll() is None:
                # print("COMSERVER is already running.")
                return
            # Run the .bat file
            self.comserver_process = subprocess.Popen(self.setup_window.comserver_path, stdout=subprocess.PIPE,
                                                      stderr=subprocess.PIPE, shell=True, text=True,
                                                      universal_newlines=True)
            # self.comserver_status_label.setText("COMSERVER Status: Running")
            self.comserver_status_label.setText("COMSERVER Status: <span style='color: green;'>Running</span>")
            self.comserver_running = True
            # Start the loop in a separate thread
            self.start_thread_to_ignore_unnecessary_lines()

        except Exception as e:
            print(f"An error occurred: {e}")

    def stop_comserver_pressed(self):
        self.comserver_running = False

    def stop_comserver(self):
        # Stop the COMSERVER process if it's running
        # self.comserver_running = False
        self.comserver_status_label.setText("COMSERVER Status: <span style='color: red;'>Not Running</span>")

        if self.comserver_process:
            try:
                # Send the SIGTERM signal to request termination
                # self.comserver_process.send_signal(signal.CTRL_C_EVENT)
                # self.comserver_process.send_signal(signal.CTRL_C_EVENT)
                # Terminate the process
                subprocess.check_call(["taskkill", "/F", "/T", "/PID", str(self.comserver_process.pid)])
                self.comserver_process = None
            except Exception as e:
                print(f"An error occurred while terminating the COMSERVER process gracefully: {e}")
                try:
                    # If SIGTERM fails, try sending a SIGKILL signal to forcefully terminate the process
                    self.comserver_process.send_signal(signal.CTRL_BREAK_EVENT)
                    # Wait for the process to terminate
                    self.comserver_process.wait()
                except Exception as e:
                    print(f"An error occurred while forcefully terminating the COMSERVER process: {e}")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    setup_window = TestSetupWindow()
    sys.exit(app.exec_())
