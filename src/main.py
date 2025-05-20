# -*- coding: utf-8 -*-
import base64
import os
import re
import sys
import pandas as pd
from PIL import Image, ImageDraw, ImageFont, ImageOps
from io import BytesIO
import time
import requests
from openai import OpenAI
from urllib.parse import quote

# Настройки
INPUT_FILE = 'articles.xlsx'  # Файл с артикулами (первый лист, второй столбец)
OUTPUT_FOLDER = 'product_cards'  # Папка для сохранения карточек
FONT_PATH = 'arial.ttf'  # Путь к шрифту (можно заменить на красивый шрифт)
PERPLEXITY_API_KEY = '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'  # Ваш API ключ Perplexity
PEXELS_API_KEY = 'LL7VOO8r9vmajcOiFTrnxUucqZO7XxC7mStcsjNCniBNaLGedWBpbPeI'  # API ключ для поиска изображений
DELAY = 2  # Задержка между запросами (секунды)
PLACEHOLDER_IMAGE = 'placeholder.jpg'


# Создаем папки
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs('fonts', exist_ok=True)
os.makedirs('temp_images', exist_ok=True)

# Цветовая схема
COLORS = {
    'background': (255, 255, 255),
    'primary': (45, 55, 72),  # Темно-синий
    'secondary': (230, 230, 230),  # Светло-серый
    'accent': (220, 60, 50),  # Красный
    'text': (60, 60, 60),  # Темно-серый
    'light_text': (150, 150, 150)  # Светло-серый
}


def create_placeholder_image():
    """Создает изображение-заглушку если его нет"""
    if not os.path.exists(PLACEHOLDER_IMAGE):
        try:
            img = Image.new('RGB', (800, 500), (245, 245, 245))
            draw = ImageDraw.Draw(img)
            font = ImageFont.truetype("fonts/NotoSans-Bold.ttf", 40)
            text = "Изображение отсутствует"
            text_width = draw.textlength(text, font=font)
            draw.text(
                ((800 - text_width) // 2, 230),
                text,
                fill=(200, 200, 200),
                font=font
            )
            img.save(PLACEHOLDER_IMAGE)
            print(f"Создано изображение-заглушка: {PLACEHOLDER_IMAGE}")
        except:
            print("Не удалось создать заглушку, будет использован белый фон")


create_placeholder_image()


def download_font(font_url, font_path):
    """Скачиваем шрифт, если его нет"""
    if not os.path.exists(font_path):
        try:
            response = requests.get(font_url)
            with open(font_path, 'wb') as f:
                f.write(response.content)
            print(f"Шрифт скачан: {font_path}")
        except Exception as e:
            print(f"Ошибка загрузки шрифта: {e}")
            return False
    return True


# Конфигурация шрифтов
FONT_CONFIG = {
    'regular': {
        'url': 'https://github.com/googlefonts/noto-fonts/raw/main/hinted/ttf/NotoSans/NotoSans-Regular.ttf',
        'path': 'fonts/NotoSans-Regular.ttf'
    },
    'bold': {
        'url': 'https://github.com/googlefonts/noto-fonts/raw/main/hinted/ttf/NotoSans/NotoSans-Bold.ttf',
        'path': 'fonts/NotoSans-Bold.ttf'
    },
    'light': {
        'url': 'https://github.com/googlefonts/noto-fonts/raw/main/hinted/ttf/NotoSans/NotoSans-Light.ttf',
        'path': 'fonts/NotoSans-Light.ttf'
    }
}

# Скачиваем шрифты
for font_type in FONT_CONFIG.values():
    if not download_font(font_type['url'], font_type['path']):
        print("Не удалось загрузить шрифты, попробуйте вручную скачать")
        sys.exit(1)


def load_font(font_type='regular', size=12):
    """Загрузка шрифта из локальной папки"""
    try:
        font_path = FONT_CONFIG[font_type]['path']
        return ImageFont.truetype(font_path, size, encoding='unic')
    except Exception as e:
        print(f"Ошибка загрузки шрифта {font_type}: {e}")
        return ImageFont.load_default(size)


# Инициализация клиента Perplexity
client = OpenAI(api_key=PERPLEXITY_API_KEY, base_url="https://api.perplexity.ai")


def sanitize_filename(filename):
    """Создание безопасного имени файла"""
    return re.sub(r'[<>:"/\\|?*]', '_', filename)


def search_product_image(article):
    """Поиск изображения через Perplexity API"""
    try:
        response = client.chat.completions.create(
            model="sonar-large-online",
            messages=[
                {
                    "role": "system",
                    "content": "Ты помощник, который находит только URL изображения товаров. Отвечай только URL изображения без каких-либо комментариев."
                },
                {
                    "role": "user",
                    "content": f"Найди URL изображения товара с артикулом {article}. Ответ должен содержать только URL изображения в формате JPG или PNG."
                }
            ]
        )

        # Извлекаем URL из ответа
        response_text = response.choices[0].message.content
        url_match = re.search(r'https?://[^\s]+\.(?:jpg|jpeg|png)', response_text)

        if url_match:
            image_url = url_match.group(0)
            response = requests.get(image_url)
            temp_path = f"temp_images/{article}_temp.jpg"
            with open(temp_path, 'wb') as f:
                f.write(response.content)
            return Image.open(temp_path)

        return Image.open(PLACEHOLDER_IMAGE) if os.path.exists(PLACEHOLDER_IMAGE) else None
    except Exception as e:
        print(f"Ошибка поиска изображения: {str(e)}")
        return Image.open(PLACEHOLDER_IMAGE) if os.path.exists(PLACEHOLDER_IMAGE) else None


def perplexity_search(article):
    """Получение информации о товаре"""
    try:
        response = client.chat.completions.create(
            model="sonar-pro",
            messages=[
                {
                    "role": "system",
                    "content": "Ты помощник, который предоставляет только технические характеристики товаров. Удаляй все сноски и квадратные скобки с цифрами из ответа."
                },
                {
                    "role": "user",
                    "content": f"Дай полное описание и все характеристики товара с артикулом {article}. Только технические характеристики, без цены, коммерческих предложений и сносок. Форматируй ответ как маркированный список, удаляя все квадратные скобки с цифрами."
                }
            ]
        )
        return response.choices[0].message.content
    except Exception as e:
        print(f"Ошибка запроса к Perplexity: {e}")
        return None


def clean_text(text):
    """Очистка текста от сносок и лишних символов"""
    if not text:
        return ""

    # Удаляем квадратные скобки с цифрами
    text = re.sub(r'\[\d+\]', '', text)
    # Удаляем упоминания цены
    text = re.sub(r'цена|стоимость|руб|₽|р\.', '', text, flags=re.IGNORECASE)
    # Удаляем лишние пробелы и переносы
    text = re.sub(r'\s+', ' ', text).strip()
    return text


def extract_product_info(article):
    """Извлечение информации о товаре"""
    search_results = perplexity_search(article)
    if not search_results:
        return None

    # Очистка текста
    cleaned_details = []
    for line in search_results.split('\n'):
        line = clean_text(line)
        if line:
            # Заменяем маркеры на единый стиль
            line = re.sub(r'^[\s•\-*]+', '▪ ', line)
            cleaned_details.append(line)

    # Поиск изображения
    image_data = search_product_image(article)

    return {
        'name': article,
        'details': cleaned_details,
        'image': image_data
    }


def create_product_card(article, product_info):
    """Создание профессиональной карточки товара"""
    try:
        # Размеры и фон
        card_width, card_height = 800, 1200
        card = Image.new('RGB', (card_width, card_height), COLORS['background'])
        draw = ImageDraw.Draw(card)

        # Загружаем шрифты
        title_font = load_font('bold', 36)
        subtitle_font = load_font('bold', 28)
        text_font = load_font('regular', 22)
        bullet_font = load_font('bold', 22)

        # Позиционирование
        y_pos = 40

        # Изображение товара (с заглушкой если нет изображения)
        img = product_info['image'] or Image.open(PLACEHOLDER_IMAGE) if os.path.exists(PLACEHOLDER_IMAGE) else None

        if img:
            # Масштабируем с сохранением пропорций
            img_ratio = img.width / img.height
            max_width = card_width - 80
            max_height = 400

            if img_ratio > 1:  # Горизонтальное
                new_width = min(max_width, img.width)
                new_height = int(new_width / img_ratio)
            else:  # Вертикальное
                new_height = min(max_height, img.height)
                new_width = int(new_height * img_ratio)

            img = img.resize((new_width, new_height), Image.LANCZOS)

            # Добавляем белую рамку
            bordered_img = ImageOps.expand(img, border=1, fill='white')
            # Добавляем тень
            shadow = Image.new('RGBA', (new_width + 6, new_height + 6), (0, 0, 0, 30))
            card.paste(shadow, (37, y_pos + 3), shadow)
            # Вставляем изображение
            x_pos = (card_width - new_width) // 2
            card.paste(bordered_img, (x_pos, y_pos))
            y_pos += new_height + 40

        # Заголовок артикула
        draw.text(
            (40, y_pos),
            f"АРТИКУЛ: {article}",
            fill=COLORS['primary'],
            font=title_font
        )
        y_pos += 50

        # Разделительная линия
        draw.line([(40, y_pos), (card_width - 40, y_pos)], fill=COLORS['secondary'], width=2)
        y_pos += 30

        # Заголовок характеристик
        draw.text(
            (40, y_pos),
            "ХАРАКТЕРИСТИКИ",
            fill=COLORS['primary'],
            font=subtitle_font
        )
        y_pos += 50

        # Характеристики с маркерами
        bullet = "▪"
        bullet_width = draw.textlength(bullet, font=bullet_font)
        max_line_width = card_width - 80 - bullet_width

        for line in product_info['details']:
            if line.strip():
                # Разбиваем длинные строки
                words = line.split()
                current_line = []

                # Первая строка с маркером
                draw.text((40, y_pos), bullet, fill=COLORS['accent'], font=bullet_font)

                for word in words:
                    test_line = ' '.join(current_line + [word])
                    if draw.textlength(test_line, font=text_font) <= max_line_width:
                        current_line.append(word)
                    else:
                        if current_line:
                            draw.text(
                                (40 + bullet_width + 10, y_pos),
                                ' '.join(current_line),
                                fill=COLORS['text'],
                                font=text_font
                            )
                            y_pos += 35
                        current_line = [word]

                if current_line:
                    draw.text(
                        (40 + bullet_width + 10, y_pos),
                        ' '.join(current_line),
                        fill=COLORS['text'],
                        font=text_font
                    )
                    y_pos += 35

                if y_pos > card_height - 50:
                    break

        # Сохранение
        safe_name = sanitize_filename(article)
        output_path = os.path.join(OUTPUT_FOLDER, f"{safe_name}.jpg")
        card.save(output_path, quality=95, optimize=True, dpi=(300, 300))
        print(f"Карточка сохранена: {output_path}")

        return True
    except Exception as e:
        print(f"Ошибка создания карточки: {e}")
        return False


def process_articles(file_path):
    """Обработка всех артикулов"""
    try:
        df = pd.read_excel(file_path)
        articles = df.iloc[:, 1].dropna().astype(str).unique()

        print(f"Найдено {len(articles)} артикулов")

        for article in articles:
            article = article.strip()
            print(f"\nОбработка: {article}")

            product_info = extract_product_info(article)
            if product_info:
                create_product_card(article, product_info)
            else:
                print(f"Не найдена информация для {article}")

            time.sleep(DELAY)

    except Exception as e:
        print(f"Ошибка обработки файла: {e}")


if __name__ == '__main__':
    print("=== Генератор профессиональных карточек товаров ===")
    print("Используются автономные шрифты Noto Sans")
    process_articles(INPUT_FILE)
    print("\n=== Обработка завершена ===")