import sys
import pandas as pd
import shutil
import os
from PyQt5 import QtWidgets, QtGui, QtCore
from PyQt5.QtWidgets import QFileDialog, QMessageBox

class ExcelProcessorApp(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Data Processing and CSV Download")

        # Calculate the center position of the screen
        screen = QtWidgets.QDesktopWidget().screenGeometry()
        width, height = screen.width(), screen.height()
        window_width, window_height = 600, 400
        self.setGeometry((width - window_width) // 2, (height - window_height) // 2, window_width, window_height)

        central_widget = QtWidgets.QWidget()
        self.setCentralWidget(central_widget)

        layout = QtWidgets.QVBoxLayout(central_widget)
        layout.setContentsMargins(20, 20, 20, 20)

        self.file_entry = QtWidgets.QLineEdit()
        self.file_entry.setFixedWidth(400)
        layout.addWidget(self.file_entry)

        self.upload_button = QtWidgets.QPushButton("Upload Excel File")
        self.upload_button.clicked.connect(self.open_file_dialog)
        self.upload_button.setStyleSheet("background-color: blue; color: white;")
        layout.addWidget(self.upload_button)

        self.execute_button = QtWidgets.QPushButton("Execute Data Processing")
        self.execute_button.clicked.connect(self.process_excel_file)
        self.execute_button.setStyleSheet("background-color: green; color: white;")
        self.execute_button.setEnabled(False)
        layout.addWidget(self.execute_button)

        self.button1 = QtWidgets.QPushButton("Download Intermittent CSV")
        self.button1.clicked.connect(self.download_intermittent_csv)
        self.button1.setEnabled(False)
        self.button1.setStyleSheet("background-color: blue; color: white;")
        layout.addWidget(self.button1)

        self.button2 = QtWidgets.QPushButton("Download Outfile CSV")
        self.button2.clicked.connect(self.download_outfile_csv)
        self.button2.setEnabled(False)
        self.button2.setStyleSheet("background-color: blue; color: white;")
        layout.addWidget(self.button2)

        self.canvas = QtWidgets.QGraphicsView()
        layout.addWidget(self.canvas)

        self.colors = ["red", "green", "blue", "orange", "purple"]
        self.circle_index = 0
        self.draw_dynamic_circle()

        # Enable file drag and drop
        self.setAcceptDrops(True)

    def open_file_dialog(self):
        file_dialog = QFileDialog()
        file_path, _ = file_dialog.getOpenFileName(self, "Upload Excel File", "", "Excel files (*.xlsx);;All files (*.*)")
        if file_path:
            self.file_entry.setText(file_path)
            self.execute_button.setEnabled(True)
            # Reset circle color to red when a new file is uploaded
            self.reset_circle_color()

    def process_excel_file(self):
        file_path = self.file_entry.text()
        try:
            data = pd.read_excel(file_path, header=None)
            melted_df = pd.melt(data, id_vars=data.columns[0], value_vars=data.columns[1:], var_name='class Name', value_name='Class')
            melted_df['class Name'] = melted_df['class Name'].astype(str).str.extract(r'(\d+)').astype(int)
            self.sorted_class_names = melted_df.groupby('Class')[0].apply(list).reset_index().sort_values(0, key=lambda x: x.str.len(), ascending=False)
            self.button1.setEnabled(True)
            self.button2.setEnabled(True)
            # Change circle color to green after data processing completion
            self.set_circle_color('green')
            QMessageBox.information(self, "Data Processing Complete", "Data processing completed successfully.")

        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred during data processing: {str(e)}")

    # def download_intermittent_csv(self):
    #     sort_class_names_csv = 'Intermittent.csv'
    #     self.sorted_class_names[['Class']].to_csv(sort_class_names_csv, index=False, header=False)
    #     QMessageBox.information(self, "Download Complete", f"Intermittent.csv downloaded successfully.")

    # def download_outfile_csv(self):
    #     non_empty_lists = self.sorted_class_names[self.sorted_class_names[0].apply(lambda x: len(x) > 0)]
    #     csv_data = pd.DataFrame(non_empty_lists[0].tolist(), index=non_empty_lists['Class'])
    #     output_csv_path = 'outfile.csv'
    #     csv_data.to_csv(output_csv_path)
    #     QMessageBox.information(self, "Download Complete", f"outfile.csv downloaded successfully.")
    
    def download_intermittent_csv(self):
        sort_class_names_csv = 'Intermittent.csv'
        save_dialog = QFileDialog()
        save_path, _ = save_dialog.getSaveFileName(self, "Save Intermittent CSV", sort_class_names_csv, "CSV files (*.csv)")
        if save_path:
            try:
                self.sorted_class_names[['Class']].to_csv(save_path, index=False, header=False)
                QMessageBox.information(self, "Download Complete", f"Intermittent.csv downloaded successfully to: {save_path}")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to save Intermittent.csv: {str(e)}")

    def download_outfile_csv(self):
        output_csv_path = 'outfile.csv'
        save_dialog = QFileDialog()
        save_path, _ = save_dialog.getSaveFileName(self, "Save Outfile CSV", output_csv_path, "CSV files (*.csv)")
        if save_path:
            try:
                non_empty_lists = self.sorted_class_names[self.sorted_class_names[0].apply(lambda x: len(x) > 0)]
                csv_data = pd.DataFrame(non_empty_lists[0].tolist(), index=non_empty_lists['Class'])
                csv_data.to_csv(save_path)
                QMessageBox.information(self, "Download Complete", f"outfile.csv downloaded successfully to: {save_path}")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to save outfile.csv: {str(e)}")


    def draw_dynamic_circle(self):
        scene = QtWidgets.QGraphicsScene()
        self.canvas.setScene(scene)
        pen = QtGui.QPen(QtGui.QColor(self.colors[self.circle_index % len(self.colors)]))
        brush = QtGui.QBrush(QtGui.QColor(self.colors[self.circle_index % len(self.colors)]))
        scene.addEllipse(10, 10, 40, 40, pen, brush)
        self.circle_index += 1
        if self.execute_button.isEnabled():
            self.start_circle_animation()

    def start_circle_animation(self):
        self.timer = QtCore.QTimer()
        self.timer.timeout.connect(self.draw_dynamic_circle)
        self.timer.start(500)  # 500 ms interval for color change

    def reset_circle_color(self):
        # Reset circle color to red
        self.set_circle_color('red')

    def set_circle_color(self, color):
        # Set circle color based on input color ('red', 'green', etc.)
        self.canvas.scene().clear()  # Clear existing scene
        pen = QtGui.QPen(QtGui.QColor(color))
        brush = QtGui.QBrush(QtGui.QColor(color))
        self.canvas.scene().addEllipse(10, 10, 40, 40, pen, brush)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        for url in event.mimeData().urls():
            file_path = url.toLocalFile()
            if file_path.lower().endswith(('.xls', '.xlsx')):
                self.file_entry.setText(file_path)
                self.execute_button.setEnabled(True)
                self.reset_circle_color()  # Reset circle color when a new file is dropped
                break

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = ExcelProcessorApp()
    window.show()
    sys.exit(app.exec_())
