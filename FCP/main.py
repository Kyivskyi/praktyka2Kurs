import sys
import subprocess
import logging
import utils.logging_config  # Імпорт модуля логування
import os
from PySide6.QtWidgets import QApplication
from ui.main_window import FileConverter
from pathlib import Path
import shutil
import urllib.request
import ctypes
import zipfile


class FFmpegManager:
    """Клас для управління перевіркою та встановленням FFmpeg."""
    
    def __init__(self, check_mode=1):
        """Ініціалізація з режимом перевірки: 0 - стандартний, 1 - тільки локальний."""
        self.check_mode = check_mode
        print(f"Ініціалізація FFmpegManager з check_mode: {self.check_mode}")
        
        # Шляхи
        self.script_dir = Path(__file__).parent
        self.ffmpeg_dir = self.script_dir / "ffmpeg"
        self.local_ffmpeg_path = self.ffmpeg_dir / "ffmpeg.exe"
        self.ffmpeg_url = "https://www.gyan.dev/ffmpeg/builds/ffmpeg-release-essentials.zip"
        self.zip_path = self.script_dir / "ffmpeg.zip"
        self.temp_dir = self.ffmpeg_dir / "temp"
        
        self.logging = logging.getLogger(__name__)
        
        # Створюємо папку ffmpeg одразу при ініціалізації
        self._create_ffmpeg_dir()

    def _create_ffmpeg_dir(self):
        """Створює папку ffmpeg, якщо її немає."""
        try:
            if not self.ffmpeg_dir.exists():
                self.ffmpeg_dir.mkdir()
                print(f"Створено папку ffmpeg: {self.ffmpeg_dir}")
                self.logging.info(f"Створено папку ffmpeg: {self.ffmpeg_dir}")
            return True
        except Exception as e:
            print(f"Помилка при створенні папки ffmpeg: {str(e)}")
            self.logging.error(f"Помилка при створенні папки ffmpeg: {str(e)}")
            return False

    def check_system_ffmpeg(self):
        """Перевіряє наявність FFmpeg у системному PATH."""
        try:
            subprocess.run(["ffmpeg", "-version"], check=True, 
                          stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            print("Загальний FFmpeg знайдено в системному PATH")
            self.logging.info("FFmpeg знайдено в системному PATH")
            return True
        except subprocess.CalledProcessError:
            print("Помилка перевірки загального FFmpeg")
            self.logging.warning("FFmpeg не знайдено в системному PATH або версія некоректна")
            return False
        except FileNotFoundError:
            print("Загальний FFmpeg не знайдено в системному PATH")
            self.logging.warning("FFmpeg не знайдено в системному PATH")
            return False

    def check_local_ffmpeg(self):
        """Перевіряє наявність FFmpeg у локальній папці."""
        if not self._create_ffmpeg_dir():  # Перевіряємо/створюємо папку перед перевіркою
            return False
            
        if self.local_ffmpeg_path.exists():
            try:
                subprocess.run([str(self.local_ffmpeg_path), "-version"], check=True,
                              stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                print("Локальний FFmpeg знайдено")
                self.logging.info("Локальний FFmpeg знайдено")
                os.environ["PATH"] = f"{self.local_ffmpeg_path.parent};{os.environ['PATH']}"
                print("Локальний FFmpeg додано до PATH")
                self.logging.info("Локальний FFmpeg додано до PATH")
                return True
            except subprocess.CalledProcessError:
                print("Помилка перевірки локального FFmpeg")
                self.logging.error("Помилка перевірки локального FFmpeg")
                return False
        print("Локальний FFmpeg не знайдено")
        self.logging.warning("Локальний FFmpeg не знайдено")
        return False

    def install_ffmpeg(self):
        """Встановлює FFmpeg у локальну папку, якщо його немає."""
        if not self._create_ffmpeg_dir():  # Перевіряємо/створюємо папку перед встановленням
            return False
            
        print("Починаю встановлення FFmpeg...")
        self.logging.info("Починаю завантаження FFmpeg...")
        try:
            # Завантаження архіву
            urllib.request.urlretrieve(self.ffmpeg_url, str(self.zip_path))

            # Тимчасове розпакування в temp
            if self.temp_dir.exists():
                shutil.rmtree(self.temp_dir)
            self.temp_dir.mkdir()
            
            with zipfile.ZipFile(self.zip_path, 'r') as zip_ref:
                zip_ref.extractall(self.temp_dir)

            # Пошук папки bin і витягування .exe
            for item in self.temp_dir.iterdir():
                if item.is_dir():
                    bin_dir = item / "bin"
                    if bin_dir.exists():
                        for bin_file in bin_dir.iterdir():
                            if bin_file.name.endswith(('.exe')):
                                target_path = self.ffmpeg_dir / bin_file.name
                                shutil.move(str(bin_file), str(target_path))

            # Видалення тимчасової папки
            shutil.rmtree(self.temp_dir)
            print("Тимчасова папка видалена")

            print("FFmpeg успішно розпаковано")
            self.logging.info("FFmpeg успішно розпаковано у локальну папку")

            # Видалення архіву
            self.zip_path.unlink()
            print("Архів FFmpeg видалено")
            self.logging.info("Архів FFmpeg видалено")

            # Перевірка після встановлення
            return self.check_local_ffmpeg()
            
        except Exception as e:
            print(f"Помилка встановлення FFmpeg: {str(e)}")
            self.logging.error(f"Помилка встановлення FFmpeg: {str(e)}")
            return False

    def ensure_ffmpeg(self):
        """Гарантує наявність FFmpeg залежно від режиму check_mode."""
        print(f"Виконання ensure_ffmpeg з check_mode: {self.check_mode}")
        
        # Спочатку перевіряємо/створюємо папку
        if not self._create_ffmpeg_dir():
            return False
            
        if self.check_mode == 0:  # Стандартний режим
            print("Перевірка загального FFmpeg...")
            if not self.check_system_ffmpeg():
                print("Перевірка локального FFmpeg...")
                if not self.check_local_ffmpeg():
                    print("Встановлення FFmpeg...")
                    return self.install_ffmpeg()
            return True
        elif self.check_mode == 1:  # Тільки локальний режим
            print("Перевірка локального FFmpeg (ігнорується системний PATH)...")
            if not self.check_local_ffmpeg():
                print("Встановлення FFmpeg у локальну папку...")
                return self.install_ffmpeg()
            return True
        return False


def main():
    """Точка входу програми."""
    print("Поточна робоча директорія:", os.getcwd())
    logging.info("Програма запущена")

    # Створення екземпляра класу з режимом перевірки (0 - стандартний, 1 - тільки локальний)
    ffmpeg_manager = FFmpegManager(check_mode=0)

    # Перевірка та встановлення FFmpeg
    if not ffmpeg_manager.ensure_ffmpeg():
        print("Не вдалося забезпечити FFmpeg.")
        ctypes.windll.user32.MessageBoxW(0, "Не вдалося забезпечити FFmpeg. Перевірте лог-файл.", "Помилка", 0x10)
        sys.exit(1)

    app = QApplication(sys.argv)
    converter = FileConverter()
    converter.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()