from PyQt5.QtWidgets import (
    QMainWindow, QSplitter, QListWidget, 
    QVBoxLayout, QWidget, QLabel
)
from PyQt5.QtCore import Qt
from raw_loader import load_sanitized_data
from DealerInfoPanel import DealerInfoPanel
from PersonnelInfoPanel import PersonnelInfoPanel
import pandas as pd


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Dealer-Personnel System")
        self.setGeometry(100, 100, 1000, 700)
        self.init_ui()
        self.load_data()
    
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
        self.course_list_widget = QListWidget()
        self.categories_list = QListWidget()
        self.categories_list.setMaximumHeight(150)  # Limit height


        right_panels = QSplitter(Qt.Vertical)
        right_panels.addWidget(self.personnel_name_label)
        right_panels.addWidget(self.personnel_position_label)
        right_panels.addWidget(self.categories_list)
        right_panels.addWidget(QLabel("دوره‌ها:"))
        right_panels.addWidget(self.course_list_widget)

        
        horizontal_splitter.addWidget(left_sidebar)
        horizontal_splitter.addWidget(right_panels)
        horizontal_splitter.setSizes([200, 600])
        
        main_layout.addWidget(horizontal_splitter)
        
        self.raw_dealer_list.currentItemChanged.connect(self.show_dealer_info)
        self.personnel_list.currentItemChanged.connect(self.show_personnel_info)
    
    def load_data(self):
        self.raw = load_sanitized_data("res/raw.xlsx")
        self.dealers = load_sanitized_data("res/dealers.xlsx")
        self.after =load_sanitized_data("res/after.xlsx")
        self.sales =load_sanitized_data("res/sales.xlsx")
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


            # Better approach to ensure uniqueness
            personnel_entries = []
            seen_combinations = set()

            for _, row in dealer_personnel.iterrows():
                name = row['نام و نام خانوادگی']
                main_pos = row.get('عنوان شغل', '')
                alt_pos = row.get('شغل موازی (ارتقا)', '')
                pcode = row.get('کد پرسنلی', '')

                # Handle position logic
                positions_to_add = []

                if pd.notna(main_pos) and main_pos.strip():
                    positions_to_add.append(main_pos.strip())
                
                if pd.notna(alt_pos) and alt_pos.strip():
                    alt_positions = [p.strip() for p in alt_pos.split('&&&') if p.strip()]
                    positions_to_add.extend(alt_positions)
                
                # If no positions found, use default
                if not positions_to_add:
                    positions_to_add = ['بدون سمت']

                # Create entries for each position
                for pos_text in positions_to_add:
                    # Create unique key using name, position, and personnel code
                    combination_key = (name, pos_text, pcode)
                    
                    if combination_key not in seen_combinations:
                        seen_combinations.add(combination_key)
                        display_text = f"{dealer_name[:4]} | {name} | {pos_text} | {pcode}"
                        personnel_entries.append(display_text)

            # Sort personnel entries before adding to list
            personnel_entries.sort(key=lambda x: x.split('|')[1].strip())  # Sort by name
            
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


    def show_personnel_info(self, item):
        if item:
            text = item.text()
            if "|" in text:
                dealer, name, position, pcode = [part.strip() for part in text.split("|")]
                # Filter relevant records
                personnel_data = self.raw[
                    self.raw['کد پرسنلی'] == pcode
                ]
                
                # Extract course names
                course_list = personnel_data['عنوان دوره'].dropna().unique().tolist()

                # Update UI directly
                self.personnel_name_label.setText(f"نام پرسنل: {name}")
                self.personnel_position_label.setText(f"سمت: {position}")
                self.course_list_widget.clear()
                self.course_list_widget.addItems(course_list)