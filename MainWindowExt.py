import webbrowser
import pandas as pd
import plotly.express as px
from PyQt6.QtWidgets import QMainWindow, QFileDialog, QMessageBox
from chartplot.MainWindow import Ui_MainWindow


class MainWindowExt(Ui_MainWindow):
    def setupUi(self, MainWindow):
        # Gọi hàm setupUi của Ui_MainWindow để thiết lập giao diện
        super().setupUi(MainWindow)
        self.MainWindow = MainWindow

        # Kết nối các nút bấm với các hàm xử lý
        self.pushButtonBrowser.clicked.connect(self.openFileDialog)
        self.pushButtonOpen.clicked.connect(self.openChartInBrowser)
        self.pushButtonSaveToHTML.clicked.connect(self.saveChartToHTML)

    def showWindow(self):
        self.MainWindow.show()

    def openFileDialog(self):
        """
        Mở hộp thoại chọn file Excel và hiển thị đường dẫn trên lineEditDataset.
        """
        file_path, _ = QFileDialog.getOpenFileName(
            self.MainWindow,
            "Open Dataset",
            "",
            "Excel Files (*.xlsx *.xls)"
        )
        if file_path:
            self.lineEditDataset.setText(file_path)

    def openChartInBrowser(self):
        """
        Đọc file Excel, vẽ biểu đồ Sunburst và mở ngay trong trình duyệt.
        """
        dataset_path = self.lineEditDataset.text().strip()
        if not dataset_path:
            QMessageBox.warning(self.MainWindow, "Error", "Please select a dataset first!")
            return

        try:
            # Đọc dữ liệu từ Excel
            df = pd.read_excel(dataset_path, sheet_name='Sheet1')
            # Làm sạch cột 'Tín Chỉ'
            df['Tín Chỉ'] = (
                df['Tín Chỉ']
                .astype(str)
                .str.extract(r'(\d+)', expand=False)
                .astype(float)
            )
            # Vẽ Sunburst Chart
            fig = px.sunburst(
                df,
                path=['Học Kỳ', 'Bắt buộc/tự chọn', 'Tên học phần'],
                values='Tín Chỉ',
                color='Học Kỳ',
                title="Chương trình học - Sunburst Chart"
            )
            # Xuất tạm ra file HTML và mở trình duyệt
            temp_html = "temp_sunburst.html"
            fig.write_html(temp_html)
            webbrowser.open(temp_html)
        except Exception as e:
            QMessageBox.critical(self.MainWindow, "Error", f"Cannot open chart.\nError details: {e}")

    def saveChartToHTML(self):
        """
        Đọc file Excel, vẽ biểu đồ Sunburst và lưu ra file HTML theo lựa chọn của người dùng.
        """
        dataset_path = self.lineEditDataset.text().strip()
        if not dataset_path:
            QMessageBox.warning(self.MainWindow, "Error", "Please select a dataset first!")
            return

        # Mở hộp thoại lưu file HTML
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
