# excel_chart_ui/ui/main_window.py

import os
from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QPushButton, QLabel,
    QComboBox, QFileDialog, QListWidget, QListWidgetItem,
    QHBoxLayout, QMessageBox
)
from PyQt6.QtCore import Qt
from utils.excel_loader import load_excel_sheets_and_columns
from charts.chart_factory import create_chart
from charts.exporter import export_chart

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Chart Generator")
        self.df = None
        self.sheets = {}

        self.open_btn = QPushButton("📂 Open Excel File")
        self.sheet_selector = QComboBox()
        self.chart_type_selector = QComboBox()
        self.x_axis_selector = QComboBox()
        self.y_axis_selector = QListWidget()
        self.generate_btn = QPushButton("📊 Generate Chart")
        self.export_btn = QPushButton("💾 Export Chart")

        self.chart_type_selector.addItems([
            "line", "bar", "column", "scatter", "area", "pie"
        ])
        self.y_axis_selector.setSelectionMode(QListWidget.SelectionMode.MultiSelection)

        layout = QVBoxLayout()
        layout.addWidget(self.open_btn)
        layout.addWidget(QLabel("Sheet:"))
        layout.addWidget(self.sheet_selector)
        layout.addWidget(QLabel("Chart Type:"))
        layout.addWidget(self.chart_type_selector)
        layout.addWidget(QLabel("X Axis:"))
        layout.addWidget(self.x_axis_selector)
        layout.addWidget(QLabel("Y Axis (multi-select supported):"))
        layout.addWidget(self.y_axis_selector)

        btns = QHBoxLayout()
        btns.addWidget(self.generate_btn)
        btns.addWidget(self.export_btn)
        layout.addLayout(btns)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

        self.open_btn.clicked.connect(self.load_excel)
        self.sheet_selector.currentIndexChanged.connect(self.update_columns)
        self.generate_btn.clicked.connect(self.render_chart)
        self.export_btn.clicked.connect(self.save_chart)

        self.current_fig = None

    def load_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)")
        if file_path:
            try:
                self.sheets = load_excel_sheets_and_columns(file_path)
                self.sheet_selector.clear()
                self.sheet_selector.addItems(self.sheets.keys())
                self.update_columns()  # trigger initial update
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to load Excel: {str(e)}")

    def update_columns(self):
        sheet = self.sheet_selector.currentText()
        if sheet in self.sheets:
            df, columns = self.sheets[sheet]
            self.df = df
            self.x_axis_selector.clear()
            self.x_axis_selector.addItems(columns)

            self.y_axis_selector.clear()
            for col in columns:
                self.y_axis_selector.addItem(QListWidgetItem(col))

    def render_chart(self):
        chart_type = self.chart_type_selector.currentText()
        x_col = self.x_axis_selector.currentText()
        y_items = self.y_axis_selector.selectedItems()
        y_cols = [item.text() for item in y_items]

        if not self.df or not x_col or not y_cols:
            QMessageBox.warning(self, "Incomplete", "Please select X and Y axis columns")
            return

        try:
            self.current_fig = create_chart(self.df, chart_type, x_col, y_cols)
            self.current_fig.show()
        except Exception as e:
            QMessageBox.critical(self, "Chart Error", str(e))

    def save_chart(self):
        if not self.current_fig:
            QMessageBox.warning(self, "No Chart", "Please generate a chart first.")
            return

        path, _ = QFileDialog.getSaveFileName(self, "Export Chart", "chart.png", "PNG Files (*.png);;PDF Files (*.pdf);;SVG Files (*.svg);;HTML Files (*.html)")
        if path:
            try:
                export_chart(self.current_fig, path)
                QMessageBox.information(self, "Success", f"Chart saved to {os.path.basename(path)}")
            except Exception as e:
                QMessageBox.critical(self, "Export Failed", str(e))
