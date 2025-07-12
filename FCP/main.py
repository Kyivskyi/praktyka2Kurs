import sys
from PySide6.QtWidgets import QApplication
from ui.main_window import FileConverter
import logging
import utils.logging_config  # Імпорт модуля логування

# Перевірка робочої директорії
import os
print("Поточна робоча директорія:", os.getcwd())

# Примусова ініціалізація логування
logging.info("Програма запущена")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    converter = FileConverter()
    converter.show()
    sys.exit(app.exec())