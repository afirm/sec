from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QTabWidget, QTableWidget,
    QTableWidgetItem, QPushButton, QHBoxLayout, 
    QHeaderView, QWidget, QLabel, QComboBox,
    QLineEdit, QHBoxLayout  # Add these imports
)
from PyQt5.QtCore import Qt
import csv
import os
import pandas as pd

class NormalizerDialog(QDialog):
    def __init__(self, parent, raw_df, dealers_df, after_sheets, sales_sheets):
        super().__init__(parent)
        self.setWindowTitle("Data Normalization Tool")
        self.setGeometry(300, 300, 1000, 700)
        
        # Store references to data
        self.raw_df = raw_df
        self.dealers_df = dealers_df
        self.after_sheets = after_sheets
        self.sales_sheets = sales_sheets
        
        layout = QVBoxLayout()
        self.tabs = QTabWidget()
        
        # Create tabs for different mappings
        self.position_tab = self.create_position_tab()
        self.car_tab = self.create_car_tab()
        self.company_tab = self.create_company_tab()
        self.course_tab = self.create_course_tab()
        
        self.tabs.addTab(self.position_tab, "Position Mappings")
        self.tabs.addTab(self.car_tab, "Car Category Mappings")
        self.tabs.addTab(self.company_tab, "Company Mappings")
        # In NormalizerDialog.__init__ after creating other tabs

        self.tabs.addTab(self.course_tab, "Course Mappings")
        
        # Buttons
        btn_layout = QHBoxLayout()
        self.save_btn = QPushButton("Save Mappings")
        self.cancel_btn = QPushButton("Cancel")
        
        self.save_btn.clicked.connect(self.save_mappings)
        self.cancel_btn.clicked.connect(self.reject)
        
        btn_layout.addWidget(self.save_btn)
        btn_layout.addWidget(self.cancel_btn)
        
        layout.addWidget(self.tabs)
        layout.addLayout(btn_layout)
        self.setLayout(layout)
        
        # Load existing mappings
        self.load_mappings()
    
    def create_position_tab(self):
        widget = QWidget()
        layout = QVBoxLayout()
        
        # Extract unique positions from raw data
        main_positions = self.raw_df['عنوان شغل'].dropna().unique().tolist()
        alt_positions = self.raw_df['شغل موازی (ارتقا)'].dropna().str.split('&&&').explode().str.strip().unique().tolist()
        all_positions = sorted(set(main_positions + alt_positions))
        
        # Extract ALL unique positions from after data
        after_positions = set()
        for sheet_name, df in self.after_sheets.items():
            if 'پست کاری' in df.columns:
                after_positions.update(df['پست کاری'].dropna().astype(str).unique())
        
        # Extract ALL unique positions from sales data
        sales_positions = set()
        for sheet_name, df in self.sales_sheets.items():
            if 'پست کاری' in df.columns:
                sales_positions.update(df['پست کاری'].dropna().astype(str).unique())
        
        # Combine all standardized positions
        all_standard_positions = sorted(after_positions.union(sales_positions))
        
        # Create table with suggestions
        table = QTableWidget()
        table.setColumnCount(3)
        table.setHorizontalHeaderLabels(["Raw Position", "Mapped Position", "Suggested Mappings"])
        table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        table.setRowCount(len(all_positions))
        
        for i, position in enumerate(all_positions):
            table.setItem(i, 0, QTableWidgetItem(position))
            table.setItem(i, 1, QTableWidgetItem(""))
            
            # Create combo box with ALL standard positions
            combo = QComboBox()
            combo.addItem("")  # Empty option
            for std_pos in all_standard_positions:
                combo.addItem(std_pos)
            
            table.setCellWidget(i, 2, combo)
        
        layout.addWidget(table)
        self.position_table = table
        widget.setLayout(layout)
        return widget
    
    def create_car_tab(self):
        widget = QWidget()
        layout = QVBoxLayout()
        
        # Extract car categories from dealers data (columns D-AV)
        car_categories = self.dealers_df.columns[3:48].tolist()
        
        # Extract ALL unique car names from after data
        after_cars = set()
        for sheet_name, df in self.after_sheets.items():
            if 'نام خودرو' in df.columns:
                after_cars.update(df['نام خودرو'].dropna().astype(str).unique())
        
        # Extract ALL unique car names from sales data
        sales_cars = set()
        for sheet_name, df in self.sales_sheets.items():
            if 'نام خودرو' in df.columns:
                sales_cars.update(df['نام خودرو'].dropna().astype(str).unique())
        
        # Combine all standardized car names
        all_standard_cars = sorted(after_cars.union(sales_cars))
        
        # Create table with suggestions
        table = QTableWidget()
        table.setColumnCount(3)
        table.setHorizontalHeaderLabels(["Raw Category", "Mapped Car", "Suggested Mappings"])
        table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        table.setRowCount(len(car_categories))
        
        for i, category in enumerate(car_categories):
            table.setItem(i, 0, QTableWidgetItem(category))
            table.setItem(i, 1, QTableWidgetItem(""))
            
            # Create combo box with ALL standard cars
            combo = QComboBox()
            combo.addItem("")  # Empty option
            for car in all_standard_cars:
                combo.addItem(car)
            
            table.setCellWidget(i, 2, combo)
        
        layout.addWidget(table)
        self.car_table = table
        widget.setLayout(layout)
        return widget
    
    def create_company_tab(self):
        widget = QWidget()
        layout = QVBoxLayout()
        
        # Extract unique companies from raw data
        companies = self.raw_df['company'].dropna().unique().tolist()
        
        # Extract sheet names from after data
        after_sheets = sorted(self.after_sheets.keys())
        
        # Extract sheet names from sales data
        sales_sheets = sorted(self.sales_sheets.keys())
        
        # Combine all sheet names
        all_sheets = sorted(set(after_sheets).union(set(sales_sheets)))
        
        # Create table with suggestions
        table = QTableWidget()
        table.setColumnCount(3)
        table.setHorizontalHeaderLabels(["Raw Company", "Mapped Company", "Suggested Mappings"])
        table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        table.setRowCount(len(companies))
        
        for i, company in enumerate(companies):
            table.setItem(i, 0, QTableWidgetItem(company))
            table.setItem(i, 1, QTableWidgetItem(""))
            
            # Create combo box with ALL sheet names
            combo = QComboBox()
            combo.addItem("")  # Empty option
            for sheet in all_sheets:
                combo.addItem(sheet)
            
            table.setCellWidget(i, 2, combo)
        
        layout.addWidget(table)
        self.company_table = table
        widget.setLayout(layout)
        return widget
    
    # Add this to your NormalizerDialog class
    
    def filter_course_table(self):
        """Filter course table based on search text"""
        search_text = self.course_search.text()
        self.populate_course_table(search_text)

    def create_course_tab(self):
        widget = QWidget()
        layout = QVBoxLayout()
        
        # Extract unique courses from raw data
        raw_courses = self.raw_df['عنوان دوره'].dropna().unique().tolist()
        
        # Extract unique courses from after data
        after_courses = set()
        for sheet_name, df in self.after_sheets.items():
            if 'نام دوره آموزشی' in df.columns:
                after_courses.update(df['نام دوره آموزشی'].dropna().astype(str).unique())
        
        # Extract unique courses from sales data
        sales_courses = set()
        for sheet_name, df in self.sales_sheets.items():
            if 'نام دوره آموزشی' in df.columns:
                sales_courses.update(df['نام دوره آموزشی'].dropna().astype(str).unique())
        
        # Combine all standardized courses
        self.all_standard_courses = sorted(after_courses.union(sales_courses))
        
        # Create search bar
        search_layout = QHBoxLayout()
        search_label = QLabel("Search Courses:")
        self.course_search = QLineEdit()
        self.course_search.setPlaceholderText("Type to filter courses...")
        self.course_search.textChanged.connect(self.filter_course_table)
        
        search_layout.addWidget(search_label)
        search_layout.addWidget(self.course_search)
        layout.addLayout(search_layout)
        
        # Create table with suggestions
        table = QTableWidget()
        table.setColumnCount(3)
        table.setHorizontalHeaderLabels(["Raw Course", "Mapped Course", "Available Standard Courses"])
        table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        table.setRowCount(len(raw_courses))
        
        self.course_table = table
        self.raw_course_list = raw_courses
        
        # Populate the table
        self.populate_course_table()
        
        layout.addWidget(table)
        widget.setLayout(layout)
        return widget

    def populate_course_table(self, filter_text=""):
        """Populate the course table with optional filtering"""
        filter_text = filter_text.lower()
        
        # Clear existing rows
        self.course_table.setRowCount(0)
        
        # Filter courses based on search text
        filtered_courses = self.raw_course_list
        if filter_text:
            filtered_courses = [course for course in self.raw_course_list 
                               if filter_text in course.lower()]
        
        # Set new row count
        self.course_table.setRowCount(len(filtered_courses))
        
        # Populate the table
        for i, course in enumerate(filtered_courses):
            self.course_table.setItem(i, 0, QTableWidgetItem(course))
            self.course_table.setItem(i, 1, QTableWidgetItem(""))
            
            # Create label with standard courses (filtered if needed)
            if filter_text:
                # Filter standard courses that match the search
                filtered_standard = [sc for sc in self.all_standard_courses 
                                    if filter_text in sc.lower()]
                standard_label = QLabel(", ".join(filtered_standard))
            else:
                standard_label = QLabel(", ".join(self.all_standard_courses))
                
            standard_label.setWordWrap(True)
            standard_label.setAlignment(Qt.AlignTop)
            self.course_table.setCellWidget(i, 2, standard_label)


    def load_mappings(self):
        # Load existing mappings if available
        self.load_mapping_file('mappings/position_mapping.csv', self.position_table)
        self.load_mapping_file('mappings/car_mapping.csv', self.car_table)
        self.load_mapping_file('mappings/company_mapping.csv', self.company_table)
        self.load_mapping_file('mappings/course_mapping.csv', self.course_table)

    
    def load_mapping_file(self, path, table):
        if os.path.exists(path):
            with open(path, 'r', encoding='utf-8') as f:
                reader = csv.reader(f)
                next(reader)  # Skip header
                
                mapping_dict = {}
                for row in reader:
                    if len(row) >= 2:
                        mapping_dict[row[0]] = row[1]
                
                # Apply mappings to table
                for row in range(table.rowCount()):
                    raw_item = table.item(row, 0)
                    if raw_item and raw_item.text() in mapping_dict:
                        # Set mapped value
                        table.setItem(row, 1, QTableWidgetItem(mapping_dict[raw_item.text()]))
                        
                        # Also select matching value in combo box if possible
                        combo = table.cellWidget(row, 2)
                        if combo:
                            index = combo.findText(mapping_dict[raw_item.text()])
                            if index >= 0:
                                combo.setCurrentIndex(index)
    
    def save_mappings(self):
        # Create mappings directory if not exists
        os.makedirs("mappings", exist_ok=True)
        
        # Save each mapping type
        self.save_mapping_type('position', self.position_table, 'position_mapping.csv')
        self.save_mapping_type('car', self.car_table, 'car_mapping.csv')
        self.save_mapping_type('company', self.company_table, 'company_mapping.csv')
        self.save_mapping_type('course', self.course_table, 'course_mapping.csv')

        
        self.accept()
    
    def save_mapping_type(self, map_type, table, filename):
        path = os.path.join("mappings", filename)
        
        with open(path, 'w', encoding='utf-8', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(["Raw", "Mapped"])
            
            for row in range(table.rowCount()):
                raw = table.item(row, 0).text() if table.item(row, 0) else ""
                
                if map_type == 'course':
                    # For course tab, mapped text is from column 1 (QTableWidgetItem)
                    mapped = table.item(row, 1).text() if table.item(row, 1) else ""
                else:
                    # For other tabs, mapped text is from combobox widget in column 2
                    combo = table.cellWidget(row, 2)
                    mapped = combo.currentText() if combo else ""
                
                if raw and mapped:
                    writer.writerow([raw, mapped])


