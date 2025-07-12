from PySide6.QtWidgets import QListWidget
from PySide6.QtGui import QDragEnterEvent, QDropEvent

class DropListWidget(QListWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.main_window = parent
        self.setAcceptDrops(True)
        self.setStyleSheet("""
            QListWidget {
                background-color: #f5f5f5;
                border: 1px solid #ccc;
                padding: 5px;
                font-size: 12px;
            }
            QListWidget::item {
                border-bottom: 1px solid #ddd;
            }
            QListWidget::item:hover {
                background-color: #e9e9e9;
            }
        """)
    
    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
    
    def dragMoveEvent(self, event: QDragEnterEvent):
        event.acceptProposedAction()
    
    def dropEvent(self, event: QDropEvent):
        files = [url.toLocalFile() for url in event.mimeData().urls()]
        if files and self.main_window:
            self.main_window.process_files(files)