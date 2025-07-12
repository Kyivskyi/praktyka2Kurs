from PySide6.QtWidgets import QWidget, QHBoxLayout, QCheckBox, QLabel, QComboBox
from utils.constants import CONVERTIBLE_FORMATS
import os

class FileItemWidget(QWidget):
    def __init__(self, file_path, parent=None):
        super().__init__(parent)
        self.file_path = file_path
        self.file_name = os.path.basename(file_path)
        self.file_ext = os.path.splitext(file_path)[1].lower()
        
        self.setup_ui()
    
    def setup_ui(self):
        layout = QHBoxLayout()
        layout.setContentsMargins(5, 5, 5, 5)
        
        self.checkbox = QCheckBox()
        self.checkbox.setMaximumWidth(20)
        self.checkbox.setChecked(False)
        self.checkbox.setVisible(True) 
        
        self.file_label = QLabel(self.file_name)
        self.file_label.setMinimumWidth(200)
        
        self.convert_label = QLabel(self.tr("→ Конвертувати в:"))
        
        self.format_combo = QComboBox()
        if self.file_ext in CONVERTIBLE_FORMATS:
            self.format_combo.addItems(CONVERTIBLE_FORMATS[self.file_ext])
        else:
            self.format_combo.addItem(self.tr("(не підтримується)"))
            self.format_combo.setEnabled(False)
        
        layout.addWidget(self.checkbox)
        layout.addWidget(self.file_label)
        layout.addWidget(self.convert_label)
        layout.addWidget(self.format_combo)
        layout.addStretch()
        
        self.setLayout(layout)