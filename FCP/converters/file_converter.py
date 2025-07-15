from PySide6.QtWidgets import QApplication, QMessageBox
from pdfminer.high_level import extract_text
from docx2pdf import convert
from pdf2docx import Converter
from PIL import Image
import pandas as pd
import json
import os
import logging
import ffmpeg
from ffmpeg._run import Error as FFmpegError

class FileConverterMixin:
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
        try:
            if not os.path.exists(pdf_path):
                raise FileNotFoundError(self.tr("PDF файл не знайдено: {}").format(pdf_path))

            cv = Converter(pdf_path)
            cv.convert(
                docx_path,
                start=0,
                end=None,
                multi_processing=True,
                recognize_tables=True,
                layout_analysis=True,
                keep_empty_lines=False,
                show_progress=False
            )
            cv.close()
            
            if not os.path.exists(docx_path):
                raise RuntimeError(self.tr("Конвертований файл не був створений: {}").format(docx_path))
                
            logging.info(f"pdf2docx: Успішно конвертовано {pdf_path} -> {docx_path}")
            
        except Exception as e:
            error_message = f"Помилка pdf2docx при конвертації {pdf_path} у {docx_path}: {str(e)}"
            logging.error(error_message, exc_info=True)
            raise Exception(self.tr("Помилка конвертації PDF у DOCX: {}. Деталі у лог-файлі.").format(str(e)))