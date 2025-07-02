from PySide6.QtWidgets import (
    QApplication, QMainWindow, QPushButton, QListWidget, QListWidgetItem,
    QFileDialog, QVBoxLayout, QWidget, QLabel, QComboBox,
    QMessageBox, QHBoxLayout, QCheckBox, QProgressBar, QGroupBox,
    QDialog, QListWidget, QVBoxLayout
)
from PySide6.QtCore import Qt, QMimeData, QTranslator, QLocale, QLibraryInfo
from PySide6.QtGui import QDragEnterEvent, QDropEvent
from pdfminer.high_level import extract_text
from docx2pdf import convert
from pdf2docx import Converter
from PIL import Image

import sys
import os
import shutil
import pandas as pd
import json
import docx
import logging
import fitz 
import ffmpeg

# Імпорт винятку Error з ffmpeg._run
from ffmpeg._run import Error as FFmpegError

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
    ".mp4": [".mkv", ".mov", ".flv", ".wmv", ".mp3"],
    ".mkv": [".mp4", ".avi", ".mov", ".flv", ".wmv"],
    ".mov": [".mp4", ".avi", ".mkv", ".flv", ".wmv"],
    ".flv": [".mp4", ".avi", ".mkv", ".mov", ".wmv"],
    ".wmv": [".mp4", ".avi", ".mkv", ".mov", ".flv"],
    ".mp3": [".wav", ".ogg"],
    ".wav": [".mp3", ".ogg"],
    ".ogg": [".mp3", ".wav"]
}

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
        
        # Встановлюємо поточний вибраний елемент
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
        self.current_language = "uk"  # Українська за замовчуванням
        # self.success_message_template = self.tr("Успішно конвертовано %1 файлів!")
        self.translator = QTranslator(self)
        
        self.load_translation()
        self.init_ui()
        self.retranslate_ui()
    
    def init_ui(self):
        self.setWindowTitle(self.tr("Конвертер файлів"))
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
        
        self.file_list = DropListWidget(self)
        self.output_label = QLabel(self.tr("Папка для збереження: Не обрано"))
        self.output_label.setStyleSheet("padding: 5px; background-color: #f0f0f0; border: 1px solid #ccc;")
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        
        self.btn_add = QPushButton(self.tr("Додати файли"))
        self.btn_convert = QPushButton(self.tr("Конвертувати всі"))
        self.btn_clear = QPushButton(self.tr("Очистити список"))
        self.btn_folder = QPushButton(self.tr("Обрати папку"))
        self.btn_delete_selected = QPushButton(self.tr("Видалити (вибрані/всі)"))
        self.btn_delete_selected.setVisible(True)
        self.btn_language = QPushButton(self.tr("Змінити мову"))
        
        self.output_folder = os.path.join(os.path.expanduser("~"), "ConvertedFiles")
        os.makedirs(self.output_folder, exist_ok=True)
        self.output_label.setText(self.tr("Папка для збереження: {}").format(self.output_folder))
        
        self.setup_ui()
        self.setup_connections()
    
    def setup_ui(self):
        central_widget = QWidget()
        layout = QVBoxLayout()

        self.button_group = QGroupBox(self.tr("Дії"))
        button_layout = QHBoxLayout()
        button_layout.addWidget(self.btn_add)
        button_layout.addWidget(self.btn_folder)
        button_layout.addWidget(self.btn_delete_selected)
        button_layout.addWidget(self.btn_language)
        self.button_group.setLayout(button_layout)

        self.file_group = QGroupBox(self.tr("Файли для конвертації"))
        file_layout = QVBoxLayout()
        file_layout.addWidget(self.file_list)
        self.file_group.setLayout(file_layout)

        layout.addWidget(self.file_group)
        layout.addWidget(self.button_group)
        layout.addWidget(self.output_label)
        layout.addWidget(self.progress_bar)
        layout.addWidget(self.btn_convert)

        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)
    
    def setup_connections(self):
        self.btn_add.clicked.connect(self.add_files)
        self.btn_folder.clicked.connect(self.select_output_folder)
        self.btn_clear.clicked.connect(self.clear_list)
        self.btn_convert.clicked.connect(self.convert_all_files)
        self.btn_delete_selected.clicked.connect(self.delete_selected_files)
        self.btn_language.clicked.connect(self.show_language_dialog)
    
    def show_language_dialog(self):
        dialog = LanguageDialog(self.current_language, self)
        if dialog.exec() == QDialog.Accepted:
            new_language = dialog.selected_language()
            if new_language != self.current_language:
                self.current_language = new_language
                self.load_translation()
                self.retranslate_ui()
    
    def load_translation(self):
        translations_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "translations")
        
        if self.current_language == "uk":
            QApplication.removeTranslator(self.translator)
        else:
            translation_file = f"fileconverter_{self.current_language}.qm"
            if self.translator.load(translation_file, translations_dir):
                QApplication.installTranslator(self.translator)
            else:
                print(f"Не вдалося завантажити переклад: {translation_file}")
    
    def retranslate_ui(self):
        # Оновлення всіх текстів інтерфейсу
        self.setWindowTitle(self.tr("Конвертер файлів"))
        self.btn_add.setText(self.tr("Додати файли"))
        self.btn_convert.setText(self.tr("Конвертувати всі"))
        self.btn_clear.setText(self.tr("Очистити список"))
        self.btn_folder.setText(self.tr("Обрати папку"))
        self.btn_delete_selected.setText(self.tr("Видалити (вибрані/всі)"))
        self.btn_language.setText(self.tr("Змінити мову"))
        self.output_label.setText(self.tr("Папка для збереження: {}").format(self.output_folder))
        self.file_group.setTitle(self.tr("Файли для конвертації"))
        self.button_group.setTitle(self.tr("Дії"))
        self.success_message_template = self.tr("Успішно конвертовано %1 файлів!")
        
        # Оновлення текстів у списку файлів
        for i in range(self.file_list.count()):
            item = self.file_list.item(i)
            widget = self.file_list.itemWidget(item)
            if widget:
                widget.convert_label.setText(self.tr("→ Конвертувати в:"))
                if not widget.format_combo.isEnabled():
                    widget.format_combo.setItemText(0, self.tr("(не підтримується)"))
    
    def add_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, self.tr("Оберіть файли"))
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
        folder = QFileDialog.getExistingDirectory(self, self.tr("Оберіть папку для збереження"))
        if folder:
            self.output_folder = folder
            self.output_label.setText(self.tr("Папка для збереження: {}").format(folder))
            os.makedirs(folder, exist_ok=True)
    
    def clear_list(self):
        if self.file_list.count() == 0:
            QMessageBox.information(self, self.tr("Очистити список"), self.tr("Список файлів вже порожній."))
            return

        reply = QMessageBox.question(self, self.tr("Підтвердження очищення"),
                                     self.tr("Ви впевнені, що хочете очистити весь список ({} файлів)?").format(self.file_list.count()),
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.file_list.clear()
            QMessageBox.information(self, self.tr("Очистити список"), self.tr("Список файлів очищено."))
            logging.info("Список файлів очищено.")
        else:
            logging.info("Очищення списку скасовано користувачем.")
    
    def delete_selected_files(self):
        items_to_delete = []
        for i in range(self.file_list.count()):
            item = self.file_list.item(i)
            widget = self.file_list.itemWidget(item)
            if widget and widget.checkbox.isChecked():
                items_to_delete.append(item)
        
        total_files_in_list = self.file_list.count()

        if items_to_delete:
            num_selected = len(items_to_delete)
            reply = QMessageBox.question(self, self.tr("Підтвердження видалення"),
                                         self.tr("Ви впевнені, що хочете видалити {} вибраних файлів зі списку?").format(num_selected),
                                         QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            
            if reply == QMessageBox.Yes:
                for item in reversed(items_to_delete):
                    row = self.file_list.row(item)
                    self.file_list.takeItem(row)
                QMessageBox.information(self, self.tr("Видалення файлів"), self.tr("Успішно видалено {} файлів зі списку.").format(num_selected))
                logging.info(f"Видалено {num_selected} вибраних файлів зі списку.")
            else:
                logging.info("Видалення вибраних файлів скасовано користувачем.")
        else:
            if total_files_in_list == 0:
                QMessageBox.information(self, self.tr("Видалення файлів"), self.tr("Список файлів порожній."))
                return

            reply = QMessageBox.question(self, self.tr("Підтвердження видалення"),
                                         self.tr("Немає вибраних файлів. Ви впевнені, що хочете видалити ВСІ {} файлів зі списку?").format(total_files_in_list),
                                         QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            
            if reply == QMessageBox.Yes:
                self.file_list.clear()
                QMessageBox.information(self, self.tr("Видалення файлів"), self.tr("Успішно видалено всі {} файлів зі списку.").format(total_files_in_list))
                logging.info(f"Видалено всі {total_files_in_list} файлів зі списку.")
            else:
                logging.info("Видалення всіх файлів скасовано користувачем.")
    
    def convert_all_files(self):
        if not hasattr(self, 'output_folder') or not self.output_folder:
            QMessageBox.warning(self, self.tr("Помилка"), self.tr("Оберіть папку для збереження!"))
            return
            
        if self.file_list.count() == 0:
            QMessageBox.warning(self, self.tr("Помилка"), self.tr("Немає файлів для конвертації!"))
            return
            
        if not os.access(self.output_folder, os.W_OK):
            QMessageBox.critical(self, self.tr("Помилка"), self.tr("Немає прав на запис у папку {}").format(self.output_folder))
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
            
            if not widget or not widget.format_combo.isEnabled():
                errors.append(f"{widget.file_name}: {self.tr('Непідтримуваний формат для конвертації')}")
                self.progress_bar.setValue(self.progress_bar.value() + 1)
                QApplication.processEvents()
                continue
                
            target_format = widget.format_combo.currentText()
            if not target_format:
                errors.append(f"{widget.file_name}: {self.tr('Не обрано цільовий формат')}")
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
                    errors.append(f"{widget.file_name}: {self.tr('Конвертований файл не створено (можлива внутрішня помилка)')}")
                    logging.error(f"Файл не створено: {output_path}")
            except Exception as e:
                errors.append(f"{widget.file_name}: {str(e)}")
                logging.error(f"Помилка конвертації {file_path}: {str(e)}")
            finally:
                self.progress_bar.setValue(self.progress_bar.value() + 1)
                QApplication.processEvents()
        
        self.progress_bar.setVisible(False)
        
        if converted > 0:
            msg = self.success_message_template.replace("%1", str(converted))
            if errors:
                msg += f"\n\n{self.error_message_template.format('\n'.join(errors))}"
            QMessageBox.information(self, self.tr("Результат"), msg)
        else:
            msg = self.tr("Жоден файл не було конвертовано!")
            if errors:
                msg += f"\n\n{self.tr('Помилки:')}\n" + "\n".join(errors)
            QMessageBox.warning(self, self.tr("Результат"), msg)
    
    def _convert_file(self, input_path, output_path, input_ext, output_ext):
        logging.info(f"Конвертація: {input_path} ({input_ext}) -> {output_path} ({output_ext})")
        
        if not os.path.exists(input_path):
            raise FileNotFoundError(self.tr("Вхідний файл не існує: {}").format(input_path))
        
        output_dir = os.path.dirname(output_path)
        os.makedirs(output_dir, exist_ok=True)
        
        if input_ext in ('.jpg', '.png', '.webp'):
            img = Image.open(input_path)
            if img.mode in ('RGBA', 'LA') and output_ext in ('.png', '.webp'):
                if output_ext == '.webp':
                    img.save(output_path, quality=90, lossless=False)
                else:
                    img.save(output_path, quality=95)
            elif output_ext == '.webp':
                if img.mode == 'RGBA':
                    img = img.convert('RGB')
                img.save(output_path, quality=90, lossless=False)
            else:
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
        
        elif input_ext == '.pdf' and output_ext == '.docx':
            self._convert_pdf_to_docx_pdf2docx(input_path, output_path)
        
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
        
        elif input_ext in ('.mp4', '.mkv', '.mov', '.flv', '.wmv') and output_ext in ('.mp4', '.mkv', '.mov', '.flv', '.wmv'):
            try:
                stream = ffmpeg.input(input_path)
                stream = ffmpeg.output(stream, output_path, vcodec='copy', acodec='copy')
                ffmpeg.run(stream, overwrite_output=True, quiet=True)
                logging.info(f"Конвертовано відео {input_ext} у {output_ext}: {output_path}")
            except ffmpeg.Error as e:
                stderr_output = e.stderr.decode('utf-8') if e.stderr else self.tr("Невідома помилка ffmpeg")
                raise Exception(self.tr("Помилка конвертації відео: {}").format(stderr_output))
        
        elif input_ext in ('.mp4', '.mkv', '.mov', '.flv', '.wmv') and output_ext == '.mp3':
            try:
                stream = ffmpeg.input(input_path)
                stream = ffmpeg.output(stream, output_path, acodec='libmp3lame')
                ffmpeg.run(stream, overwrite_output=True, quiet=True)
                logging.info(f"Вирізано звук з {input_ext} і збережено як {output_ext}: {output_path}")
            except ffmpeg.Error as e:
                stderr_output = e.stderr.decode('utf-8') if e.stderr else self.tr("Невідома помилка ffmpeg")
                raise Exception(self.tr("Помилка вирізання звуку: {}").format(stderr_output))
        
        elif input_ext in ('.mp3', '.wav', '.ogg') and output_ext in ('.mp3', '.wav', '.ogg'):
            try:
                stream = ffmpeg.input(input_path)
                if output_ext == '.ogg':
                    stream = ffmpeg.output(stream, output_path, acodec='libvorbis')
                elif output_ext == '.mp3':
                    stream = ffmpeg.output(stream, output_path, acodec='libmp3lame')
                else:  # .wav
                    stream = ffmpeg.output(stream, output_path)
                ffmpeg.run(stream, overwrite_output=True, quiet=True)
                logging.info(f"Конвертовано аудіо {input_ext} у {output_ext}: {output_path}")
            except ffmpeg.Error as e:
                stderr_output = e.stderr.decode('utf-8') if e.stderr else self.tr("Невідома помилка ffmpeg")
                raise Exception(self.tr("Помилка конвертації аудіо: {}").format(stderr_output))
        
        else:
            raise ValueError(self.tr("Непідтримувана комбінація конвертації: {} у {}").format(input_ext, output_ext))
    
    def _convert_pdf_to_docx_pdf2docx(self, pdf_path, docx_path):
        """Конвертація PDF у DOCX з використанням pdf2docx з розширеними налаштуваннями"""
        try:
            # Перевірка наявності файлу
            if not os.path.exists(pdf_path):
                raise FileNotFoundError(self.tr("PDF файл не знайдено: {}").format(pdf_path))

            # Створення конвертера з розширеними параметрами
            cv = Converter(pdf_path)
            
            # Конвертація з оптимізацією для складних PDF
            cv.convert(
                docx_path,
                start=0,                  # Почати з першої сторінки
                end=None,                 # Конвертувати до кінця документа
                multi_processing=True,    # Використовувати багатопотоковість
                recognize_tables=True,    # Краще розпізнавання таблиць
                layout_analysis=True,     # Аналіз структури (для стовпців)
                keep_empty_lines=False,   # Ігнорувати порожні рядки
                show_progress=False       # Не показувати прогрес у консолі
            )
            cv.close()
            
            # Перевірка результату
            if not os.path.exists(docx_path):
                raise RuntimeError(self.tr("Конвертований файл не був створений: {}").format(docx_path))
                
            logging.info(f"pdf2docx: Успішно конвертовано {pdf_path} -> {docx_path}")
            
        except Exception as e:
            error_message = f"Помилка pdf2docx при конвертації {pdf_path} у {docx_path}: {str(e)}"
            logging.error(error_message, exc_info=True)
            raise Exception(self.tr("Помилка конвертації PDF у DOCX: {}. Деталі у лог-файлі.").format(str(e)))


if __name__ == "__main__":
    app = QApplication(sys.argv)
    converter = FileConverter()
    converter.show()
    sys.exit(app.exec())