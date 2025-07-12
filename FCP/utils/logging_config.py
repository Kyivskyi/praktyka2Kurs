import logging
import os

project_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
log_file = os.path.join(project_dir, 'converter.log')

os.makedirs(os.path.dirname(log_file), exist_ok=True)

for handler in logging.root.handlers[:]:
    logging.root.removeHandler(handler)

logging.basicConfig(
    filename=log_file,
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    filemode='a'  # Додаємо, а не перезаписуємо
)

logging.info("Логування ініціалізовано в %s", log_file)