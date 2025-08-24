# -*- coding: utf-8 -*-
# Імпорт необхідних бібліотек
import os
import threading
import sys
import time
import win32clipboard
from pystray import Icon as TrayIcon, Menu as TrayMenu, MenuItem as TrayMenuItem
from PIL import Image, ImageDraw, ImageFont
from PIL import UnidentifiedImageError
from io import BytesIO
from huggingface_hub import snapshot_download
from transformers import TrOCRProcessor, VisionEncoderDecoderModel
from loguru import logger

# --- Нова функція для обробки шляхів до ресурсів у PyInstaller ---
def resource_path(relative_path):
    """
    Повертає абсолютний шлях до ресурсу, незалежно від того,
    чи програма запущена як скрипт, чи скомпільована за допомогою PyInstaller.
    """
    try:
        # У PyInstaller файли знаходяться у тимчасовій теці `_MEIPASS`
        base_path = sys._MEIPASS
    except Exception:
        # При запуску як звичайного скрипту, використовуємо поточний каталог
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# Встановлюємо змінну середовища, щоб Hugging Face не показував попередження про symlinks.
os.environ['HF_HUB_DISABLE_SYMLINKS_WARNING'] = '1'

# Визначаємо шлях до нашої портативної кеш-папки.
# !!! ТУТ ЗМІНИ !!!
portable_cache_dir = os.path.join(resource_path('.'), 'model_cache')
model_name = 'kha-white/manga-ocr-base'

# Глобальні змінні для Manga-OCR, потоку моніторингу та іконки в треї
mocr = None
ocr_thread = None
stop_event = threading.Event()
last_clipboard_hash = None
mocr_running = True # Додано для відстеження стану

# Визначаємо словник для заміни символів
PUNCTUATION_MAP = {
    '?': '？',
    '!': '！',
    ',': '、',
    '~': '～',
    ':': '：',
    ';': '；',
    '“': '〝',
    '”': '〟',
    '[': '【',
    ']': '】',
    '(': '（',
    ')': '）',
    '.': '・',
    '…': '・・・',
    '-': 'ー',
    '~': '〜'
}

def get_image_from_clipboard():
    """
    Отримує зображення з буфера обміну з повторними спробами.
    """
    image = None
    max_retries = 5
    retries = 0
    clipboard_open = False
    
    while retries < max_retries:
        try:
            win32clipboard.OpenClipboard()
            clipboard_open = True
            if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_DIB):
                data = win32clipboard.GetClipboardData(win32clipboard.CF_DIB)
                image_stream = BytesIO(data)
                image = Image.open(image_stream)
            break  # Успішно отримали доступ, виходимо з циклу
        except Exception as e:
            # Можливі помилки, якщо буфер обміну використовується іншим процесом
            logger.warning(f"Помилка при роботі з буфером обміну, спроба {retries + 1}/{max_retries}: {e}")
            retries += 1
            time.sleep(0.1) # Коротка пауза перед наступною спробою
        finally:
            if clipboard_open:
                win32clipboard.CloseClipboard()
    
    return image

def set_text_to_clipboard(text):
    """Поміщає текст у буфер обміну."""
    try:
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardText(text, win32clipboard.CF_UNICODETEXT)
    except Exception as e:
        # Без логування помилка не буде виведена
        pass
    finally:
        win32clipboard.CloseClipboard()

def replace_punctuation(text):
    """
    Замінює деякі символи на їх повноширинні аналоги.
    """
    for old, new in PUNCTUATION_MAP.items():
        text = text.replace(old, new)
    return text

def clean_text(text):
    """
    Видаляє всі пробіли з розпізнаного тексту, замінює пунктуацію
    і перетворює лапки на японські.
    """
    # Спочатку замінюємо пунктуацію
    text = replace_punctuation(text)
    
    # Видаляємо зайві пробіли
    cleaned_text = text.replace(' ', '')
    
    # Заміна лапок на японські. Використовуємо прапорець для відстеження
    # відкритих/закритих лапок.
    result_text = ""
    is_open_quote = True
    for char in cleaned_text:
        if char == '"':
            if is_open_quote:
                result_text += "〝"
            else:
                result_text += "〟"
            is_open_quote = not is_open_quote
        else:
            result_text += char
            
    return result_text

def monitor_clipboard():
    """Потік для моніторингу буфера обміну."""
    global last_clipboard_hash
    global mocr_running
    while not stop_event.is_set():
        # Перевірка, чи активний моніторинг
        if mocr_running:
            img = get_image_from_clipboard()
            if img:
                # Створюємо хеш зображення, щоб уникнути повторної обробки
                img_hash = hash(img.tobytes())
                if img_hash != last_clipboard_hash:
                    last_clipboard_hash = img_hash
                    try:
                        text = mocr(img)
                        # Очищаємо текст від зайвих відступів
                        cleaned_text = clean_text(text)
                        set_text_to_clipboard(cleaned_text)
                    except Exception as e:
                        # Без логування помилка не буде виведена
                        pass
        time.sleep(0.5)

def toggle_ocr(icon, item):
    """Вмикає або вимикає моніторинг."""
    global mocr_running
    mocr_running = not mocr_running
    # Залишаємо логіку перемикання стану без виведення повідомлень.

def on_exit(icon, item):
    """Функція для коректного виходу з програми."""
    global ocr_thread
    global mocr
    stop_event.set()
    if ocr_thread and ocr_thread.is_alive():
        ocr_thread.join()
    icon.stop()
    os._exit(0)

def load_model_and_start_ocr(icon):
    """
    Завантажує модель та ініціалізує OCR в окремому потоці,
    щоб не блокувати головний цикл іконки в треї.
    """
    global mocr
    global ocr_thread
    
    # --- Завантаження та ініціалізація моделі ---
    try:
        # 1. Завантажуємо модель в нашу портативну папку.
        model_path = snapshot_download(repo_id=model_name, cache_dir=portable_cache_dir, local_files_only=False)
        
        # 2. Завантажуємо процесор і модель з локальної папки
        use_fast_processor = False
        try:
            import torchvision
            use_fast_processor = True
        except ImportError:
            print("\nПопередження: Бібліотека `torchvision` не знайдена.")
            print("Для використання швидкого процесора зображень, будь ласка, встановіть її.")
            print("Виконайте наступну команду у вашому терміналі:")
            print(f".\\python.exe -m pip install torchvision")
            print("Програма продовжить роботу з повільним процесором.")

        # Ініціалізуємо процесор і модель, залежно від наявності `torchvision`.
        processor = TrOCRProcessor.from_pretrained(model_path, local_files_only=True, use_fast=use_fast_processor)
        model = VisionEncoderDecoderModel.from_pretrained(model_path, local_files_only=True)
        
        # 3. Створюємо об'єкт, який імітує поведінку MangaOcr
        class PortableMangaOcr:
            def __init__(self, processor, model):
                self.processor = processor
                self.model = model
            
            def __call__(self, image):
                pixel_values = self.processor(image, return_tensors="pt").pixel_values
                generated_ids = self.model.generate(pixel_values)
                generated_text = self.processor.batch_decode(generated_ids, skip_special_tokens=True)[0]
                return generated_text

        mocr = PortableMangaOcr(processor, model)
        
    except Exception as e:
        # Якщо сталася критична помилка, програма просто завершиться
        sys.exit(1)

    # Запуск потоку моніторингу буфера обміну після завантаження моделі
    ocr_thread = threading.Thread(target=monitor_clipboard, daemon=True)
    ocr_thread.start()
    
    # --- Завантаження фінальної іконки та повідомлення ---
    # !!! ТУТ ЗМІНИ !!!
    # Використовуємо нову функцію resource_path для пошуку іконки
    icon_path = resource_path('Favico.ico')
    try:
        final_image = Image.open(icon_path)
    except FileNotFoundError:
        # Якщо файл не знайдено, створюємо тимчасову іконку
        final_image = Image.new('RGB', (64, 64), color='gray') 
    except Exception as e:
        final_image = Image.new('RGB', (64, 64), color='gray')
    
    # Оновлюємо іконку та заголовок у треї
    icon.icon = final_image
    icon.title = 'Manga OCR'
    
    # Надсилаємо сповіщення, щоб привернути увагу
    try:
        icon.notify("Програма готова до роботи!", "Manga OCR")
    except Exception as e:
        # Іноді сповіщення можуть не працювати, тому додаємо обробку помилки
        pass


def main():
    """
    Основна функція, яка запускає іконку в треї.
    """
    
    # --- Створення тимчасової іконки для завантаження ---
    loading_icon_image = Image.new('RGB', (64, 64), color='white')
    d = ImageDraw.Draw(loading_icon_image)
    try:
        font = ImageFont.truetype("arial.ttf", 30)
    except IOError:
        font = ImageFont.load_default()
    d.text((5, 10), "...", fill='black', font=font)
    
    # Створення меню та запуск іконки в системному треї
    menu = TrayMenu(
        TrayMenuItem('Зупинити/Запустити OCR', toggle_ocr),
        TrayMenuItem('Вихід', on_exit)
    )
    
    # Створюємо іконку з індикатором завантаження
    icon = TrayIcon('manga-ocr', loading_icon_image, 'Manga OCR (Завантаження...)', menu)
    
    # Запускаємо завантаження моделі в окремому потоці
    loading_thread = threading.Thread(target=load_model_and_start_ocr, args=(icon,), daemon=True)
    loading_thread.start()

    # Головний потік запускає іконку в треї, що блокує його виконання
    # і тримає програму запущеною, поки користувач її не закриє.
    icon.run()

if __name__ == '__main__':
    main()
