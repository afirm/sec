from PyQt5.QtWidgets import (
    QMainWindow, QSplitter, QListWidget, 
    QVBoxLayout, QWidget, QLabel, QPushButton,QDialog, QScrollArea,QListWidgetItem
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor

from raw_loader import load_sanitized_data, load_all_sanitized_sheets
from DealerInfoPanel import DealerInfoPanel
import pandas as pd
from NormalizerDialog import NormalizerDialog
import os,csv
from collections import defaultdict


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Dealer-Personnel System")
        self.setGeometry(100, 100, 1000, 700)
        self.init_ui()
        self.load_data()
        self.load_mappings()  # Load mappings after data
        self.personnel_training_status = {}  


    
    def init_ui(self):
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout()
        main_widget.setLayout(main_layout)
        
        horizontal_splitter = QSplitter(Qt.Horizontal)
        
        left_sidebar = QSplitter(Qt.Vertical)
        
        self.raw_dealer_list = QListWidget()
        self.personnel_list = QListWidget()
        left_sidebar.addWidget(self.raw_dealer_list)
        left_sidebar.addWidget(self.personnel_list)
        


        self.dealer_info = DealerInfoPanel()

        self.personnel_name_label = QLabel("نام پرسنل:")
        self.personnel_position_label = QLabel("سمت:")
        # self.course_list_widget = QListWidget()


        # Replace personnel_details_label creation with:
        self.personnel_details_label = QLabel()
        self.personnel_details_label.setWordWrap(True)
        self.personnel_details_label.setTextFormat(Qt.RichText)
        # Enable text selection
        self.personnel_details_label.setTextInteractionFlags(
            Qt.TextSelectableByMouse |  # Allow mouse selection
            Qt.LinksAccessibleByMouse   # Allow clicking links
        )
        
        # Improve selection visibility
        self.personnel_details_label.setStyleSheet(
            "QLabel::selected {"
            "   background-color: #3399ff;"
            "   color: white;"
            "}"
        )


        # self.personnel_details_label = QLabel()
        # self.personnel_details_label.setWordWrap(True)
        # self.personnel_details_label.setTextFormat(Qt.RichText)

        # self.personnel_details_label = QLabel()
        # self.personnel_details_label.setWordWrap(True)
        # self.personnel_details_label.setTextFormat(Qt.RichText)

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setWidget(self.personnel_details_label)



        self.categories_list = QListWidget()
        self.categories_list.setMaximumHeight(150)  # Limit height


        right_panels = QSplitter(Qt.Vertical)
        # right_panels.addWidget(self.personnel_name_label)
        # right_panels.addWidget(self.personnel_position_label)
        # right_panels.addWidget(self.categories_list)
        # right_panels.addWidget(QLabel("دوره‌ها:"))
        # right_panels.addWidget(self.course_list_widget)
        # Instead of course_list_widget, add personnel_details_label to right_panels
        # right_panels.addWidget(self.personnel_details_label)
        right_panels.addWidget(scroll_area)
        
        horizontal_splitter.addWidget(left_sidebar)
        horizontal_splitter.addWidget(right_panels)
        horizontal_splitter.setSizes([200, 600])
        
        main_layout.addWidget(horizontal_splitter)
        
        self.raw_dealer_list.currentItemChanged.connect(self.show_dealer_info)
        self.personnel_list.currentItemChanged.connect(self.show_personnel_info)


        menubar = self.menuBar()
        settings_menu = menubar.addMenu('Settings')
        mapping_action = settings_menu.addAction('Data Normalization')
        mapping_action.triggered.connect(self.open_normalizer)
        
    def open_normalizer(self):
        # Pass our data to the normalizer dialog
        dialog = NormalizerDialog(
            self, 
            self.raw, 
            self.dealers, 
            self.after_sheets, 
            self.sales_sheets
        )
        if dialog.exec_() == QDialog.Accepted:
            self.load_mappings()  # Reload mappings if saved
    
    def load_data(self):
        self.raw = load_sanitized_data("res/raw.xlsx")
        self.dealers = load_sanitized_data("res/dealers.xlsx")
        self.after =load_sanitized_data("res/after.xlsx")
        self.sales =load_sanitized_data("res/sales.xlsx")
        # Load all worksheets from after and sales files
        self.after_sheets = load_all_sanitized_sheets("res/after.xlsx")
        self.sales_sheets = load_all_sanitized_sheets("res/sales.xlsx")

        dealers = sorted(self.raw['عنوان نمایندگی'].unique())
        self.raw_dealer_list.addItems(dealers)
          
    def show_dealer_info(self, item):
        if item:
            dealer_name = item.text()
            self.dealer_info.display_info(dealer_name, self.raw)

            # Clear and populate categories list
            self.categories_list.clear()
            categories = self.get_dealer_categories(dealer_name)
            for category in categories:
                self.categories_list.addItem(category)

            dealer_personnel = self.raw[self.raw['عنوان نمایندگی'] == dealer_name]
            self.personnel_list.clear()

            personnel_entries = []
            seen_combinations = set()
            
            # Clear previous stored training status
            self.personnel_training_status = {}

            for _, row in dealer_personnel.iterrows():
                name = row['نام و نام خانوادگی']
                main_pos = row.get('عنوان شغل', '')
                alt_pos = row.get('شغل موازی (ارتقا)', '')
                pcode = row.get('کد پرسنلی', '')
                rawcompany = row.get('company')

                # Handle position logic
                positions_to_add = []

                if pd.notna(main_pos) and main_pos.strip():
                    positions_to_add.append(main_pos.strip())
                
                if pd.notna(alt_pos) and alt_pos.strip():
                    alt_positions = [p.strip() for p in alt_pos.split('&&&') if p.strip()]
                    positions_to_add.extend(alt_positions)
                
                if not positions_to_add:
                    positions_to_add = ['بدون سمت']

                for pos_text in positions_to_add:
                    combination_key = (name, pos_text, pcode)
                    if combination_key not in seen_combinations:
                        seen_combinations.add(combination_key)

                        mapped_company = self.company_mapping.get(rawcompany, rawcompany)
                        mapped_position = self.position_mapping.get(pos_text, pos_text)
                        mapped_categories = [self.car_mapping.get(cat, cat) for cat in categories]

                        # Get required trainings from after sheets
                        after_requirements = self.get_matching_after_rows(
                            mapped_company,
                            mapped_position,
                            mapped_categories
                        )

                        # Get required trainings from sales sheets
                        sales_requirements = self.get_matching_sales_rows(
                            mapped_company,
                            mapped_position
                        )

                        # Get passed courses for this personnel
                        passed_courses = self.raw[
                            (self.raw['کد پرسنلی'] == pcode)
                        ]['عنوان دوره'].dropna().unique().tolist()
                        taken_set = set(passed_courses)

                        # Compare and tag results
                        results = []

                        for req in after_requirements:
                            criteria = req.get("نام سرفصل", "").strip()
                            car = req.get("نام خودرو", "").strip() or "عمومی"
                            course = req.get("نام دوره آموزشی", "").strip()
                            is_taken = course in taken_set

                            results.append({
                                "file": "after",
                                "criteria": criteria,
                                "car": car,
                                "course": course,
                                "is_taken": is_taken
                            })

                        for req in sales_requirements:
                            criteria = req.get("معیار", "").strip()
                            car = "فروش"
                            course = req.get("نام دوره", "").strip()
                            is_taken = course in taken_set

                            results.append({
                                "file": "sales",
                                "criteria": criteria,
                                "car": car,
                                "course": course,
                                "is_taken": is_taken
                            })


                        self.personnel_training_status[(pcode, mapped_position)] = results

                        display_text = f"{dealer_name[:4]} | {name} | {pos_text} | {pcode}"

                        if pos_text not in self.position_mapping:
                            # Create QListWidgetItem disabled (not selectable)
                            item = QListWidgetItem(display_text)
                            item.setFlags(Qt.ItemIsEnabled)  # Enabled but NOT selectable (remove ItemIsSelectable)
                            item.setForeground(QColor('gray'))  # visually gray out
                            self.personnel_list.addItem(item)
                        else:
                            # Normal selectable item
                            item = QListWidgetItem(display_text)
                            self.personnel_list.addItem(item)
                        personnel_entries.append(display_text)

            # Sort personnel entries by name
            personnel_entries.sort(key=lambda x: x.split('|')[1].strip())

            for entry in personnel_entries:
                self.personnel_list.addItem(entry)


    def get_dealer_categories(self, dealer_name):
        """Extract categories for a specific dealer from dealers data"""
        categories = []
        
        # Find the dealer row in dealers dataframe
        dealer_row = self.dealers[self.dealers.iloc[:, 0] == dealer_name[:4]]  # Column A is dealer code
        if not dealer_row.empty:
            # Get the first matching row
            row = dealer_row.iloc[0]
            
            # Check columns D to AV (columns 3 to 47 in 0-based indexing)
            # Column D is index 3, Column AV is index 47
            for col_idx in range(3, min(48, len(row))):
                if col_idx < len(self.dealers.columns):
                    category_name = self.dealers.columns[col_idx]
                    cell_value = row.iloc[col_idx]
                    
                    # Check if cell contains 'p' (case insensitive)
                    if pd.notna(cell_value) and str(cell_value).strip().lower() == 'p':
                        categories.append(category_name)
        
        return categories

    def get_matching_sales_rows(self, company, position):
        """
        Find matching rows from 'sales_sheets' for given company and position.
        Sales sheets only use 'فروش' as car category by default.
        """
        matched_rows = []
        sheet_df = self.sales_sheets.get(company)
        if sheet_df is None:
            return []

        for _, row in sheet_df.iterrows():
            row_pos = str(row.get("پست کاری", "")).strip()
            if row_pos == position:
                matched_rows.append(row)

        print(f"[SALES MATCH] company={company}, position={position} → matched {len(matched_rows)} rows")
        return matched_rows


    def get_matching_after_rows(self, company, position, cars):
        """
        Find matching rows from 'after_sheets' for given company, position, and list of cars (categories).
        If any car is empty, treat it as 'عمومی'.
        """
        if not isinstance(cars, list):
            cars = [cars]

        # Clean car names, replace empty with "عمومی"
        cars = [str(c).strip() if str(c).strip() else "عمومی" for c in cars]
        cars.append("عمومی")  # Always allow default fallback
        cars = list(set(cars))  # Ensure uniqueness

        matched_rows = []

        sheet_df = self.after_sheets.get(company)
        if sheet_df is None:
            return []
        # if sheet_df is not None:
        #     print("Sheet columns:", sheet_df.columns.tolist())
        #     print(sheet_df[['نام خودرو', 'پست کاری']].dropna().head(10).to_string())


        for _, row in sheet_df.iterrows():
            row_car = str(row.get("نام خودرو", "")).strip() or "عمومی"
            row_pos = str(row.get("پست کاری", "")).strip()

            if row_car in cars and row_pos == position:
                matched_rows.append(row)

        # print(f"[AFTER MATCH] company={company}, position={position}, cars={cars} → matched {len(matched_rows)} rows")
        return matched_rows

    def load_mappings(self):
        self.position_mapping = {}
        self.car_mapping = {}
        self.company_mapping = {}
        
        # Load position mappings
        self.load_mapping_file('mappings/position_mapping.csv', self.position_mapping)
        # Load car category mappings
        self.load_mapping_file('mappings/car_mapping.csv', self.car_mapping)
        # Load company mappings
        self.load_mapping_file('mappings/company_mapping.csv', self.company_mapping)
        self.course_mapping = {}
        self.load_mapping_file('mappings/course_mapping.csv', self.course_mapping)

    
    def load_mapping_file(self, path, mapping_dict):
        if os.path.exists(path):
            with open(path, 'r', encoding='utf-8') as f:
                reader = csv.reader(f)
                next(reader)  # Skip header
                for row in reader:
                    if len(row) >= 2:
                        mapping_dict[row[0]] = row[1]

    def show_personnel_info(self, item):
        if not item:
            return
        if not item.flags() & Qt.ItemIsSelectable:
        # Ignore disabled item
            return
        
        text = item.text()
        if "|" not in text:
            return
        
        dealer_code, name, position, pcode = [part.strip() for part in text.split("|")]
        mapped_position = self.position_mapping.get(position, position)
        key = (pcode, mapped_position)

        # Get training status results for this personnel
        results = self.personnel_training_status.get(key, [])

        # Find full dealer name and company
        dealer_full_name = next((d for d in self.raw['عنوان نمایندگی'].unique() if d.startswith(dealer_code)), "")
        filtered = self.raw[(self.raw['کد پرسنلی'] == pcode) & (self.raw['عنوان نمایندگی'] == dealer_full_name)]
        company = filtered.iloc[0].get('company', '') if not filtered.empty else ""

        passed_courses = filtered['عنوان دوره'].dropna().unique().tolist()
        passed_courses = [self.course_mapping.get(c, c) for c in passed_courses]
        passed_set = set(passed_courses)

        def colored_courses(courses, passed_set):
            parts = []
            for c in courses:
                # Skip empty or invalid courses
                if not c or str(c).strip() == "" or str(c).strip() == "nan":
                    continue
                if c in passed_set:
                    parts.append(f'<span style="background-color:#a2c1d5;">{c}</span>')  # light green
                else:
                    parts.append(f'<span style="background-color:#d3d3d3;">{c}</span>')  # grey
            return " &mdash; ".join(parts)

        passed_courses_str = colored_courses(passed_courses, passed_set)

        # Group results by file -> car -> criteria
        grouped = defaultdict(lambda: defaultdict(lambda: defaultdict(list)))
        for r in results:
            file_key = r.get("file", "after")
            grouped[file_key][r["car"]][r["criteria"]].append(r["course"])

        # Add sales courses for this person from sales_sheets
        mapped_company = self.company_mapping.get(company, company)
        sales_df = self.sales_sheets.get(mapped_company)
        if sales_df is not None:
            for _, row in sales_df.iterrows():
                row_pos = str(row.get("پست کاری", "")).strip()
                if row_pos == mapped_position:
                    # Swap the columns - A is actually course, C is actually criteria
                    course = str(row.get("نام دوره آموزشی", "")).strip()  # Column A
                    criteria = str(row.get("نام سرفصل", "")).strip()      # Column C
                    car = "فروش"
                    
                    # Skip empty or invalid criteria/courses
                    if criteria and criteria != "nan" and course and course != "nan":
                        grouped["sales"][car][criteria].append(course)

        # --- New: Determine criteria passed or not ---
        # Criteria is passed if:
        # 1) at least one course taken for it
        # 2) OR criteria includes "گازسوز"
        # 3) OR criteria includes "ابزار مخصوص" and all other criteria in the same car category are passed
        
        # We'll create a dict: pass_status[file][car][criteria] = True/False
        pass_status = defaultdict(lambda: defaultdict(dict))

        for file_name, cars in grouped.items():
            for car, criteria_dict in cars.items():
                # First pass: criteria passed if one course taken or name contains گازسوز
                for crit, courses in criteria_dict.items():
                    passed = False
                    if any(c in passed_set for c in courses):
                        passed = True
                    elif "گازسوز" in crit:
                        passed = True
                    pass_status[file_name][car][crit] = passed

                # Second pass: handle "ابزار مخصوص"
                # If criteria contains "ابزار مخصوص", it passes only if ALL OTHER criteria in that car are passed
                for crit in criteria_dict.keys():
                    if "ابزار مخصوص" in crit:
                        # Check if all other criteria in this car are passed
                        others = [c for c in criteria_dict.keys() if c != crit]
                        if all(pass_status[file_name][car].get(c, False) for c in others):
                            pass_status[file_name][car][crit] = True
                        else:
                            pass_status[file_name][car][crit] = False

        # --- Calculate pass counts and percentages per car and overall ---
        # count total and passed criteria per car and overall
        overall_total = 0
        overall_passed = 0

        criteria_html_parts = []
        for file_name, cars in grouped.items():
            file_label = "فروش" if file_name == "sales" else "خدمات پس از فروش"
            criteria_html_parts.append(f'<h1>{file_label}:</h1><br>')

            for car, crits in cars.items():
                total_criteria = len(crits)
                passed_criteria = sum(1 for crit in crits if pass_status[file_name][car].get(crit, False))

                overall_total += total_criteria
                overall_passed += passed_criteria

                percent = (passed_criteria / total_criteria * 100) if total_criteria else 0

                # Car header with passed/total and percent
                criteria_html_parts.append(
                    f'<div style="margin-right:20px;"><b>خودرو: {car} - '
                    f'{passed_criteria} از {total_criteria} (٪{percent:.1f})</b></div><br>'
                )
            
                for crit, courses in crits.items():
                    # Skip empty criteria or courses
                    if not crit or not courses or all(not c for c in courses):
                        continue
                        
                    passed = pass_status[file_name][car].get(crit, False)
                    bgcolor = "#a8d5a2" if passed else "#f8a5a5"  # light green or light red
                    course_html = colored_courses(courses, passed_set)
                    
                    # Skip if course_html is empty or only contains dashes/spaces
                    if not course_html.strip() or course_html.strip() == "&mdash;":
                        continue
                        
                    criteria_html_parts.append(
                        f'<div style="margin-right:40px; background-color:{bgcolor}; padding:4px; border-radius:4px; margin-bottom:3px;">'
                        f'{crit} ({course_html})</div>'
                    )
                criteria_html_parts.append('<br>')

        overall_percent = (overall_passed / overall_total * 100) if overall_total else 0

        # Add total pass info at top
        top_info = (
            f'<b>نام پرسنل:</b> {name} &mdash; '
            f'موارد گذرانده شده: {overall_passed} از {overall_total} (٪{overall_percent:.1f})<br>'
            f'<b>سمت:</b> {position}<br>'
            f'<b>نمایندگی:</b> {dealer_full_name}<br>'
            f'<b>شرکت:</b> {company}<br><br>'
        )

        criteria_section = "".join(criteria_html_parts) if criteria_html_parts else "— موردی یافت نشد —"

        html = (
            top_info +
            f'<b>دوره‌های الزامی و معیارها:</b><br>{criteria_section}<br><br>'
            f'<b>دوره‌های گذرانده شده:</b><br>{passed_courses_str}'
        )
        
        self.personnel_details_label.setText(html)