import sys
import pandas as pd
import numpy as np
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QLabel, QTableWidget, QTableWidgetItem
)
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure

class MarksAnalyzerApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Student Marks Analyzer")
        self.setGeometry(100, 100, 900, 600)

        # Layout
        self.layout = QVBoxLayout()

        # File Upload Buttons
        self.label = QLabel("Upload the Marks and Student Details files:")
        self.layout.addWidget(self.label)

        self.upload_marks_btn = QPushButton("Upload Marks File")
        self.upload_marks_btn.clicked.connect(self.load_marks_file)
        self.layout.addWidget(self.upload_marks_btn)

        self.upload_details_btn = QPushButton("Upload Student Details File")
        self.upload_details_btn.clicked.connect(self.load_details_file)
        self.layout.addWidget(self.upload_details_btn)

        # Analyze Button
        self.analyze_btn = QPushButton("Analyze Distribution")
        self.analyze_btn.clicked.connect(self.analyze_distribution)
        self.analyze_btn.setEnabled(False)
        self.layout.addWidget(self.analyze_btn)

        # Chart Area
        self.figure = Figure()
        self.canvas = FigureCanvas(self.figure)
        self.layout.addWidget(self.canvas)

        # Table for Data Display
        self.table = QTableWidget()
        self.layout.addWidget(self.table)

        # Set Layout
        self.setLayout(self.layout)

        # Data Storage
        self.marks_data = None
        self.details_data = None

    def load_marks_file(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "Open Marks File", "", "Excel Files (*.xlsx)")
        if file_name:
            self.marks_data = pd.read_excel(file_name)
            self.label.setText("Marks file loaded successfully!")
            self.check_ready()

    def load_details_file(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "Open Student Details File", "", "Excel Files (*.xlsx)")
        if file_name:
            self.details_data = pd.read_excel(file_name)
            self.label.setText("Student details file loaded successfully!")
            self.check_ready()


    def generate_report(self):
        # Check if files are loaded
        if self.marks_data is None or self.details_data is None:
            self.label.setText("Please upload both files first.")
            return

        # Merge data
        merged_data = pd.merge(self.marks_data, self.details_data, left_on='Symbol/Reg', right_on='Symbol Number')

        # Extract subjects and grades
        subjects = ['BIOLOGY (TH)', 'CHEMISTRY (TH)', 'COM.ENGLISH (TH)', 'COM.MATHEMATICS (TH)', 
                    'COM.NEPALI (TH)', 'COMPUTER SCIENCE', 'PHYSICS (TH)']
        report = []

        for subject in subjects:
            subject_data = merged_data[merged_data['SUB'] == subject]
            grades = subject_data['GPA']
            grade_counts = subject_data['SEE Grade'].value_counts()

            # Calculate average GPA
            avg_gpa = np.round(grades.mean(), 2)

            # Calculate percentage of valid grades
            valid_grades = grade_counts.sum() - grade_counts.get('NG', 0) - grade_counts.get('Abs', 0)
            percentage = np.round((valid_grades / grade_counts.sum()) * 100, 2) if grade_counts.sum() > 0 else 0

            # Prepare data for the report
            for grade, count in grade_counts.items():
                report.append({
                    'Subject': subject,
                    'Grade': grade,
                    'No. of Students': count,
                    'Avg': avg_gpa,
                    'Percentage': f"{percentage}%" if grade in grade_counts else "N/A"
                })

        # Convert to DataFrame for display
        report_df = pd.DataFrame(report)
        self.display_table(report_df)


    def check_ready(self):
        if self.marks_data is not None and self.details_data is not None:
            self.analyze_btn.setEnabled(True)

    def analyze_distribution(self):
        # Merge Data
        merged_data = pd.merge(self.marks_data, self.details_data, left_on='Symbol/Reg', right_on='Symbol Number')

        # Calculate Grade Distribution
        grade_counts = merged_data['SEE Grade'].value_counts()

        # Plot Grade Distribution
        self.figure.clear()
        ax = self.figure.add_subplot(111)
        ax.bar(grade_counts.index, grade_counts.values, color='skyblue')
        ax.set_title("Grade Distribution")
        ax.set_xlabel("Grades")
        ax.set_ylabel("Number of Students")
        self.canvas.draw()

        # Display Table
        self.display_table(merged_data[['Name', 'SEE Grade', 'GPA']])

    def display_table(self, data):
        self.table.setRowCount(len(data))
        self.table.setColumnCount(len(data.columns))
        self.table.setHorizontalHeaderLabels(data.columns)

        for i, row in data.iterrows():
            for j, value in enumerate(row):
                self.table.setItem(i, j, QTableWidgetItem(str(value)))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MarksAnalyzerApp()
    window.show()
    sys.exit(app.exec_())

