# personnel_info_panel.py
from PyQt5.QtWidgets import QTextEdit, QWidget, QVBoxLayout

class PersonnelInfoPanel(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.init_ui()
    
    def init_ui(self):
        layout = QVBoxLayout()
        self.text_edit = QTextEdit()
        QTextEdit().setReadOnly(True)
        layout.addWidget(self.text_edit)
        self.setLayout(layout)
    
    def display_info(self, personnel_name, df):
        personnel_data = df[df['نام و نام خانوادگی'] == personnel_name]
        info_text = f"<h2>Personnel: {personnel_name}</h2>"
        info_text += personnel_data.to_html(index=False)
        self.personnel_info.setHtml(info_text)

