from PySide6.QtWidgets import (
    QApplication, QMainWindow, QPushButton, QListWidget, QListWidgetItem,
    QFileDialog, QVBoxLayout, QWidget, QLabel, QComboBox,
    QMessageBox, QHBoxLayout, QCheckBox, QProgressBar, QGroupBox
)
from PySide6.QtCore import Qt, QMimeData
from PySide6.QtGui import QDragEnterEvent, QDropEvent
import sys
import os
import shutil
from PIL import Image
import pandas as pd
import json
import docx
from pdfminer.high_level import extract_text
from docx2pdf import convert
import logging
import fitz

# Імпортуємо Aspose.PDF
import aspose.pdf as apdf

# Налаштування логування
logging.basicConfig(filename='converter.log', level=logging.DEBUG,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Словник підтримуваних форматів конвертації
CONVERTIBLE_FORMATS = {
    ".docx": [".pdf"],
    ".pdf": [".txt", ".docx"],
    ".txt": [".pdf", ".docx"],
    ".jpg": [".png", ".webp"],
    ".png": [".jpg", ".webp"],
    ".csv": [".xlsx", ".json"],
    ".json": [".csv", ".xlsx"],
}

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
        self.checkbox.setChecked(False) # ЗМІНА: За замовчуванням чекбокс НЕ відмічений при додаванні
        self.checkbox.setVisible(True) 
        
        self.file_label = QLabel(self.file_name)
        self.file_label.setMinimumWidth(200)
        
        self.convert_label = QLabel("→ Конвертувати в:")
        
        self.format_combo = QComboBox()
        if self.file_ext in CONVERTIBLE_FORMATS:
            self.format_combo.addItems(CONVERTIBLE_FORMATS[self.file_ext])
        else:
            self.format_combo.addItem("(не підтримується)")
            self.format_combo.setEnabled(False)
        
        layout.addWidget(self.checkbox)
        layout.addWidget(self.file_label)
        layout.addWidget(self.convert_label)
        layout.addWidget(self.format_combo)
        layout.addStretch()
        
        self.setLayout(layout)

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

class FileConverter(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Конвертер файлів")
        self.setMinimumSize(800, 600)
        self.setStyleSheet("""
            QMainWindow {
                background-color: #e0e0e0;
            }
            QWidget {
                font-size: 12px;
            }
            QPushButton {
                background-color: #d0d0d0;
                padding: 6px 10px;
                border: 1px solid #aaa;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #c0c0c0;
            }
            QComboBox {
                padding: 3px;
                min-width: 100px;
            }
            QProgressBar {
                border: 1px solid #aaa;
                border-radius: 3px;
                text-align: center;
            }
            QGroupBox {
                border: 1px solid #aaa;
                border-radius: 5px;
                margin-top: 10px;
                padding-top: 15px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 3px;
            }
        """)
        
        # Головні елементи
        self.file_list = DropListWidget(self)
        self.output_label = QLabel("Папка для збереження: Не обрано")
        self.output_label.setStyleSheet("padding: 5px; background-color: #f0f0f0; border: 1px solid #ccc;")
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        
        # Кнопки
        self.btn_add = QPushButton("Додати файли")
        self.btn_convert = QPushButton("Конвертувати всі")
        self.btn_folder = QPushButton("Обрати папку")
        self.btn_delete_selected = QPushButton("Видалити (вибрані/всі)") # ЗМІНА: Оновлено текст кнопки
        self.btn_delete_selected.setVisible(True)

        # Ініціалізуємо self.output_folder за замовчуванням
        self.output_folder = os.path.join(os.path.expanduser("~"), "ConvertedFiles")
        os.makedirs(self.output_folder, exist_ok=True)
        self.output_label.setText(f"Папка для збереження: {self.output_folder}")
        
        # Налаштування інтерфейсу
        self.setup_ui()
        self.setup_connections()
    
    def setup_ui(self):
        central_widget = QWidget()
        layout = QVBoxLayout()
        
        # Група для кнопок
        button_group = QGroupBox("Дії")
        button_layout = QHBoxLayout()
        button_layout.addWidget(self.btn_add)
        button_layout.addWidget(self.btn_folder)
        button_layout.addWidget(self.btn_delete_selected) 
        button_group.setLayout(button_layout)
        
        # Група для списку файлів
        file_group = QGroupBox("Файли для конвертації")
        file_layout = QVBoxLayout()
        file_layout.addWidget(self.file_list)
        file_group.setLayout(file_layout)
        
        # Додавання елементів до основного лейауту
        layout.addWidget(file_group)
        layout.addWidget(button_group)
        layout.addWidget(self.output_label)
        layout.addWidget(self.progress_bar)
        layout.addWidget(self.btn_convert)
        
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)
    
    def setup_connections(self):
        self.btn_add.clicked.connect(self.add_files)
        self.btn_folder.clicked.connect(self.select_output_folder)
        self.btn_convert.clicked.connect(self.convert_all_files)
        self.btn_delete_selected.clicked.connect(self.delete_selected_files)
    
    def add_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Оберіть файли")
        if files:
            self.process_files(files)
    
    def process_files(self, files):
        for file in files:
            if os.path.isfile(file):
                existing_files = [self.file_list.itemWidget(self.file_list.item(i)).file_path 
                                  for i in range(self.file_list.count()) 
                                  if self.file_list.itemWidget(self.file_list.item(i))]
                if file not in existing_files:
                    item = QListWidgetItem(self.file_list)
                    widget = FileItemWidget(file)
                    item.setSizeHint(widget.sizeHint())
                    self.file_list.addItem(item)
                    self.file_list.setItemWidget(item, widget)
                else:
                    logging.info(f"Файл вже у списку: {file}")
    
    def select_output_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Оберіть папку для збереження")
        if folder:
            self.output_folder = folder
            self.output_label.setText(f"Папка для збереження: {folder}")
            os.makedirs(self.output_folder, exist_ok=True)


# минула функція яка видаляла файли по кнопці очистити список, і використовувалась як видалення після конвертації
#
#   def clear_list(self):
#        # Очищає весь список файлів без умов.
#        if self.file_list.count() == 0:
#            QMessageBox.information(self, "Очистити список", "Список файлів вже порожній.")
#            return
#
#        reply = QMessageBox.question(self, "Підтвердження очищення",
#                                     f"Ви впевнені, що хочете очистити весь список ({self.file_list.count()} файлів)?",
#                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
#        if reply == QMessageBox.Yes:
#            self.file_list.clear()
#            QMessageBox.information(self, "Очистити список", "Список файлів очищено.")
#            logging.info("Список файлів очищено.")
#        else:
#            logging.info("Очищення списку скасовано користувачем.")



    def delete_selected_files(self):
        """
        Видаляє вибрані файли. Якщо жоден файл не вибрано, питає про видалення всіх.
        """
        items_to_delete = []
        for i in range(self.file_list.count()):
            item = self.file_list.item(i)
            widget = self.file_list.itemWidget(item)
            if widget and widget.checkbox.isChecked():
                items_to_delete.append(item)
        
        total_files_in_list = self.file_list.count()

        if items_to_delete: # Якщо є вибрані файли
            num_selected = len(items_to_delete)
            reply = QMessageBox.question(self, "Підтвердження видалення",
                                         f"Ви впевнені, що хочете видалити {num_selected} вибраних файлів зі списку?",
                                         QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            
            if reply == QMessageBox.Yes:
                # Видаляємо елементи в зворотному порядку, щоб уникнути проблем з індексами
                for item in reversed(items_to_delete):
                    row = self.file_list.row(item)
                    self.file_list.takeItem(row)
                QMessageBox.information(self, "Видалення файлів", f"Успішно видалено {num_selected} файлів зі списку.")
                logging.info(f"Видалено {num_selected} вибраних файлів зі списку.")
            else:
                logging.info("Видалення вибраних файлів скасовано користувачем.")
        else: # Якщо жоден файл не вибрано
            if total_files_in_list == 0:
                QMessageBox.information(self, "Видалення файлів", "Список файлів порожній.")
                return

            reply = QMessageBox.question(self, "Підтвердження видалення",
                                         f"Немає вибраних файлів. Ви впевнені, що хочете видалити ВСІ {total_files_in_list} файлів зі списку?",
                                         QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            
            if reply == QMessageBox.Yes:
                self.file_list.clear()
                QMessageBox.information(self, "Видалення файлів", f"Успішно видалено всі {total_files_in_list} файлів зі списку.")
                logging.info(f"Видалено всі {total_files_in_list} файлів зі списку.")
            else:
                logging.info("Видалення всіх файлів скасовано користувачем.")

    def convert_all_files(self):
        if not hasattr(self, 'output_folder') or not self.output_folder:
            QMessageBox.warning(self, "Помилка", "Оберіть папку для збереження!")
            return
            
        if self.file_list.count() == 0:
            QMessageBox.warning(self, "Помилка", "Немає файлів для конвертації!")
            return
            
        if not os.access(self.output_folder, os.W_OK):
            QMessageBox.critical(self, "Помилка", f"Немає прав на запис у папку {self.output_folder}")
            return

        self.progress_bar.setVisible(True)
        
        total_files = self.file_list.count()
        
        self.progress_bar.setMaximum(total_files)
        self.progress_bar.setValue(0)
        QApplication.processEvents()
        
        converted = 0
        errors = []
        
        for i in range(self.file_list.count()):
            item = self.file_list.item(i)
            widget = self.file_list.itemWidget(item)
            
            # ЗМІНА: Тепер конвертуються ВСІ файли, незалежно від стану чекбокса
            # Чекбокси тільки для видалення
            
            if not widget or not widget.format_combo.isEnabled():
                errors.append(f"{widget.file_name}: Непідтримуваний формат для конвертації")
                self.progress_bar.setValue(self.progress_bar.value() + 1)
                QApplication.processEvents()
                continue
                
            target_format = widget.format_combo.currentText()
            if not target_format:
                errors.append(f"{widget.file_name}: Не обрано цільовий формат")
                self.progress_bar.setValue(self.progress_bar.value() + 1)
                QApplication.processEvents()
                continue
                
            file_path = widget.file_path
            file_name = os.path.splitext(os.path.basename(file_path))[0]
            output_path = os.path.join(self.output_folder, f"{file_name}{target_format}")
            
            try:
                logging.info(f"Починаємо конвертацію: {file_path} -> {output_path}")
                self._convert_file(file_path, output_path, widget.file_ext, target_format)
                if os.path.exists(output_path):
                    converted += 1
                    logging.info(f"Успішно конвертовано: {output_path}")
                else:
                    errors.append(f"{widget.file_name}: Конвертований файл не створено (можлива внутрішня помилка)")
                    logging.error(f"Файл не створено: {output_path}")
            except Exception as e:
                errors.append(f"{widget.file_name}: {str(e)}")
                logging.error(f"Помилка конвертації {file_path}: {str(e)}")
            finally:
                self.progress_bar.setValue(self.progress_bar.value() + 1)
                QApplication.processEvents()
        
        self.progress_bar.setVisible(False)
        
        if converted > 0:
            msg = f"Успішно конвертовано {converted} файлів!"
            if errors:
                msg += f"\n\nПомилки:\n" + "\n".join(errors)
            QMessageBox.information(self, "Результат", msg)
        else:
            msg = "Жоден файл не було конвертовано!"
            if errors:
                msg += f"\n\nПомилки:\n" + "\n".join(errors)
            QMessageBox.warning(self, "Результат", msg)
            
        self.file_list.clear() # Авто видалення після конвертації
    
    def _convert_file(self, input_path, output_path, input_ext, output_ext):
        logging.info(f"Конвертація: {input_path} ({input_ext}) -> {output_path} ({output_ext})")
        
        if not os.path.exists(input_path):
            raise FileNotFoundError(f"Вхідний файл не існує: {input_path}")
        
        output_dir = os.path.dirname(output_path)
        os.makedirs(output_dir, exist_ok=True)
        
        if input_ext in ('.jpg', '.png', '.webp'):
            img = Image.open(input_path)
            # Перевіряємо, чи зображення має альфа-канал і чи цільовий формат його підтримує
            if img.mode in ('RGBA', 'LA') and output_ext in ('.png', '.webp'):
                # Зберігаємо з прозорістю
                if output_ext == '.webp':
                    img.save(output_path, quality=90, lossless=False) # lossless=False для зменшення розміру з втратами
                else:
                    img.save(output_path, quality=95) # PNG зберігає якість
            elif output_ext == '.webp': # Якщо немає прозорості, але зберігаємо в WebP
                # Конвертуємо до RGB, якщо є альфа, щоб уникнути помилок при збереженні без альфа-каналу
                if img.mode == 'RGBA':
                    img = img.convert('RGB')
                img.save(output_path, quality=90, lossless=False)
            else: # Звичайне збереження без прозорості
                # Конвертуємо до RGB, якщо є альфа, щоб уникнути помилок при збереженні без альфа-каналу
                if img.mode == 'RGBA':
                    img = img.convert('RGB')
                img.save(output_path, quality=95)
            logging.info(f"Зображення збережено: {output_path}")
        
        elif input_ext == '.docx' and output_ext == '.pdf':
            convert(input_path, output_path)
            logging.info(f"Конвертовано docx у pdf: {output_path}")
        
        elif input_ext == '.pdf' and output_ext == '.txt':
            text = extract_text(input_path)
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(text)
            logging.info(f"Конвертовано pdf у txt: {output_path}")
        
        # ОНОВЛЕНО ЗГІДНО З ДОКУМЕНТАЦІЄЮ: Використовуємо DocSaveOptions.DocFormat.DOC_X
        elif input_ext == '.pdf' and output_ext == '.docx':
            self._convert_pdf_to_docx_aspose(input_path, output_path)
            logging.info(f"Конвертовано pdf у docx за допомогою Aspose.PDF: {output_path}")
        
        elif input_ext == '.csv' and output_ext == '.xlsx':
            df = pd.read_csv(input_path)
            df.to_excel(output_path, index=False, engine='openpyxl')
            logging.info(f"Конвертовано csv у xlsx: {output_path}")
        
        elif input_ext == '.csv' and output_ext == '.json':
            df = pd.read_csv(input_path)
            df.to_json(output_path, orient='records', indent=2, force_ascii=False)
            logging.info(f"Конвертовано csv у json: {output_path}")
        
        elif input_ext == '.json' and output_ext == '.csv':
            with open(input_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            df = pd.json_normalize(data)
            df.to_csv(output_path, index=False, encoding='utf-8')
            logging.info(f"Конвертовано json у csv: {output_path}")
        
        else:
            raise ValueError(f"Непідтримувана комбінація конвертації: {input_ext} у {output_ext}")

    def _convert_pdf_to_docx_aspose(self, pdf_path, docx_path):
        """Конвертація PDF у DOCX з використанням Aspose.PDF та DocSaveOptions."""
        try:
            document = apdf.Document(pdf_path)
            
            save_options = apdf.DocSaveOptions()
            
            # ВИКОРИСТОВУЄМО ПРАВИЛЬНИЙ ФОРМАТ: DOC_X згідно з наданою документацією
            save_options.format = apdf.DocSaveOptions.DocFormat.DOC_X
            
            document.save(docx_path, save_options)
            
            logging.info(f"Aspose.PDF: Успішно збережено {pdf_path} як {docx_path}")
        except Exception as e:
            error_message = f"Помилка Aspose.PDF при конвертації {pdf_path} у {docx_path}: {str(e)}"
            logging.error(error_message, exc_info=True)
            raise Exception(f"Помилка конвертації PDF у DOCX за допомогою Aspose.PDF: {str(e)}. Деталі в лог-файлі.")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    converter = FileConverter()
    converter.show()
    sys.exit(app.exec())