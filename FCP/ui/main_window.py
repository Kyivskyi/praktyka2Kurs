from PySide6.QtWidgets import (
    QApplication, QMainWindow, QPushButton, QListWidget, QListWidgetItem,
    QFileDialog, QVBoxLayout, QWidget, QLabel, QComboBox,
    QMessageBox, QHBoxLayout, QCheckBox, QProgressBar, QGroupBox,
    QDialog, QListWidget, QVBoxLayout
)
from PySide6.QtCore import QTranslator, QLocale, QLibraryInfo
from ui.file_item_widget import FileItemWidget
from ui.drop_list_widget import DropListWidget
from ui.language_dialog import LanguageDialog
from converters.file_converter import FileConverterMixin
from utils.constants import CONVERTIBLE_FORMATS
import os
import logging

class FileConverter(QMainWindow, FileConverterMixin):
    def __init__(self):
        super().__init__()
        self.current_language = "uk"  # Ukrainian by default
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
        translations_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "../translations")
        
        if self.current_language == "uk":
            QApplication.removeTranslator(self.translator)
        else:
            translation_file = f"fileconverter_{self.current_language}.qm"
            if self.translator.load(translation_file, translations_dir):
                QApplication.installTranslator(self.translator)
            else:
                logging.warning(f"Не вдалося завантажити переклад: {translation_file}")
    
    def retranslate_ui(self):
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
        self.error_message_template = self.tr("Помилки:\n%1")
        
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