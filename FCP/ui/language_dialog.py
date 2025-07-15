from PySide6.QtWidgets import QDialog, QVBoxLayout, QListWidget, QPushButton

class LanguageDialog(QDialog):
    def __init__(self, current_language, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Language")
        self.setFixedSize(300, 200)
        
        self.languages = {
            "Українська": "uk",
            "English": "en",
            "Dansk": "da",
            "Deutsch": "de",
            "Español": "es",
            "Français": "fr",
            "Italiano": "it",
            "Nederlands": "nl",
            "Polski": "pl",
            "Português": "pt",
            "Suomi": "fi",
            "Svenska": "sv",
            "Türkçe": "tr",
            "한국어": "ko",
            "日本語": "ja",
            "简体中文": "zh_CN",
        }
        
        layout = QVBoxLayout()
        
        self.list_widget = QListWidget()
        self.list_widget.addItems(self.languages.keys())
        
        for i, (name, code) in enumerate(self.languages.items()):
            if code == current_language:
                self.list_widget.setCurrentRow(i)
                break
        
        self.confirm_button = QPushButton("Підтвердити")
        self.confirm_button.clicked.connect(self.accept)
        
        layout.addWidget(self.list_widget)
        layout.addWidget(self.confirm_button)
        self.setLayout(layout)
    
    def selected_language(self):
        current_item = self.list_widget.currentItem()
        if current_item:
            return self.languages[current_item.text()]
        return "uk"