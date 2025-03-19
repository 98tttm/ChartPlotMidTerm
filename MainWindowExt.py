import webbrowser
import pandas as pd
import plotly.express as px
from PyQt6.QtWidgets import QMainWindow, QFileDialog, QMessageBox
from chartplot.MainWindow import Ui_MainWindow


class MainWindowExt(Ui_MainWindow):
    def setupUi(self, MainWindow):
        super().setupUi(MainWindow)
        self.MainWindow = MainWindow

        self.pushButtonBrowser.clicked.connect(self.openFileDialog)
        self.pushButtonOpen.clicked.connect(self.openChartInBrowser)
        self.pushButtonSaveToHTML.clicked.connect(self.saveChartToHTML)

    def showWindow(self):
        self.MainWindow.show()

    def openFileDialog(self):

        file_path, _ = QFileDialog.getOpenFileName(
            self.MainWindow,
            "Open Dataset",
            "",
            "Excel Files (*.xlsx *.xls)"
        )
        if file_path:
            self.lineEditDataset.setText(file_path)

    def openChartInBrowser(self):
        dataset_path = self.lineEditDataset.text().strip()
        if not dataset_path:
            QMessageBox.warning(self.MainWindow, "Error", "Please select a dataset first!")
            return

        try:
            df = pd.read_excel(dataset_path, sheet_name='Sheet1')
            df['Tín Chỉ'] = (
                df['Tín Chỉ']
                .astype(str)
                .str.extract(r'(\d+)', expand=False)
                .astype(float)
            )

            fig = px.sunburst(
                df,
                path=['Học Kỳ', 'Bắt buộc/tự chọn', 'Tên học phần'],
                values='Tín Chỉ',
                color='Học Kỳ',
                title="Chương trình học - Sunburst Chart"
            )
    
            temp_html = "temp_sunburst.html"
            fig.write_html(temp_html)
            webbrowser.open(temp_html)
        except Exception as e:
            QMessageBox.critical(self.MainWindow, "Error", f"Cannot open chart.\nError details: {e}")

    def saveChartToHTML(self):
        dataset_path = self.lineEditDataset.text().strip()
        if not dataset_path:
            QMessageBox.warning(self.MainWindow, "Error", "Please select a dataset first!")
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self.MainWindow,
            "Save HTML Chart",
            "chart.html",
            "HTML Files (*.html)"
        )
        if not file_path:
            return

        try:
            df = pd.read_excel(dataset_path, sheet_name='Sheet1')
            df['Tín Chỉ'] = (
                df['Tín Chỉ']
                .astype(str)
                .str.extract(r'(\d+)', expand=False)
                .astype(float)
            )
            fig = px.sunburst(
                df,
                path=['Học Kỳ', 'Bắt buộc/tự chọn', 'Tên học phần'],
                values='Tín Chỉ',
                color='Học Kỳ',
                title="Chương trình học - Sunburst Chart"
            )
            fig.write_html(file_path)
            QMessageBox.information(self.MainWindow, "Success", f"Chart saved to:\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self.MainWindow, "Error", f"Cannot save chart.\nError details: {e}")
