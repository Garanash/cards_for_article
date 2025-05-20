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
from bs4 import BeautifulSoup
from urllib.parse import quote

# Настройки
INPUT_FILE = 'articles.xlsx'  # Файл с артикулами (первый лист, второй столбец)
OUTPUT_FOLDER = 'product_cards'  # Папка для сохранения карточек
FONT_PATH = 'arial.ttf'  # Путь к шрифту (можно заменить на красивый шрифт)
PERPLEXITY_API_KEY = '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'  # Ваш API ключ Perplexity
PEXELS_API_KEY = 'LL7VOO8r9vmajcOiFTrnxUucqZO7XxC7mStcsjNCniBNaLGedWBpbPeI'  # API ключ для поиска изображений
DELAY = 2  # Задержка между запросами (секунды)
PLACEHOLDER_IMAGE = 'placeholder.png'
YANDEX_SEARCH_URL = "https://yandex.eu/images/search"  # Европейский домен

# Создаем папки
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs('fonts', exist_ok=True)
os.makedirs('temp_images', exist_ok=True)

# Цветовая схема
COLORS = {
    'background': (255, 255, 255),
    'header': (13, 71, 161),
    'footer': (13, 71, 161),
    'accent': (255, 87, 34),
    'text': (33, 33, 33),
    'light_text': (250, 250, 250)
}


def create_placeholder_image():
    """Создает изображение-заглушку если его нет"""
    if not os.path.exists(PLACEHOLDER_IMAGE):
        try:
            img = Image.new('RGB', (350, 250), (245, 245, 245))
            draw = ImageDraw.Draw(img)
            font = ImageFont.truetype("fonts/NotoSans-Bold.ttf", 20)
            text = "Изображение отсутствует"
            text_width = draw.textlength(text, font=font)
            draw.text(
                ((350 - text_width) // 2, 115),
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


def search_yandex_images(query, result_position=5):
    """Поиск изображений через европейский домен Яндекса (6-е по популярности)"""
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
            'Accept-Language': 'en-US,en;q=0.9'
        }
        params = {
            'text': query,
            'nomisspell': 1,
            'noreask': 1,
            'isize': 'medium',  # Средний размер изображений
            'itype': 'jpg'  # Предпочтительный формат
        }

        response = requests.get(YANDEX_SEARCH_URL, headers=headers, params=params, timeout=15)
        response.raise_for_status()

        soup = BeautifulSoup(response.text, 'html.parser')
        images = soup.find_all('img', class_='serp-item__thumb')

        if len(images) > result_position:
            img_url = images[result_position].get('src')
            if img_url and img_url.startswith('//'):
                img_url = 'https:' + img_url

            if img_url:
                img_response = requests.get(img_url, headers=headers, timeout=15)
                img_response.raise_for_status()

                temp_path = f"temp_images/{query}_temp.jpg"
                with open(temp_path, 'wb') as f:
                    f.write(img_response.content)

                img = Image.open(temp_path)
                if img.mode in ('RGBA', 'P'):
                    img = img.convert('RGB')
                return img

        return Image.open(PLACEHOLDER_IMAGE) if os.path.exists(PLACEHOLDER_IMAGE) else None

    except Exception as e:
        print(f"Ошибка поиска изображения в Яндекс: {str(e)}")
        return Image.open(PLACEHOLDER_IMAGE) if os.path.exists(PLACEHOLDER_IMAGE) else None


def perplexity_search(article):
    """Получение информации о товаре в формате списка характеристик"""
    try:
        response = client.chat.completions.create(
            model="sonar-pro",
            messages=[
                {
                    "role": "system",
                    "content": "Ты помощник, который предоставляет технические характеристики товаров в формате: 'Параметр: значение'. Удаляй все сноски, символы (#, *, []) и коммерческую информацию."
                },
                {
                    "role": "user",
                    "content": f"Дай технические характеристики товара с артикулом {article} в строгом формате 'Параметр: значение'. Только факты, без цен, предложений купить и комментариев. Пример:\nМатериал: сталь\nВес: 1.5 кг\nЦвет: серебристый"
                }
            ],
            temperature=0.2  # Минимум креативности для точности
        )

        # Очистка и форматирование текста
        text = response.choices[0].message.content
        lines = []
        for line in text.split('\n'):
            line = re.sub(r'[#*\[\]]', '', line.strip())  # Удаляем спецсимволы
            if line and ':' in line and not any(word in line.lower() for word in ['цена', 'стоимость', 'руб']):
                lines.append(line)

        return '\n'.join(lines)
    except Exception as e:
        print(f"Ошибка запроса к Perplexity: {e}")
        return None


def extract_product_info(article):
    """Извлечение информации о товаре"""
    search_results = perplexity_search(article)
    if not search_results:
        return None

    # Разбиваем на список характеристик
    characteristics = []
    for line in search_results.split('\n'):
        if ':' in line:
            param, value = line.split(':', 1)
            characteristics.append(f"{param.strip()}: {value.strip()}")

    # Поиск изображения через Яндекс (6-е по популярности)
    image_data = search_yandex_images(article, 5)

    return {
        'name': article,
        'characteristics': characteristics,
        'image': image_data
    }


def create_product_card(article, product_info):
    """Создание карточки товара в альбомной ориентации"""
    try:
        # Размеры и фон (альбомная ориентация)
        card_width, card_height = 1200, 800
        card = Image.new('RGB', (card_width, card_height), COLORS['background'])
        draw = ImageDraw.Draw(card)

        # Загружаем шрифты
        title_font = load_font('bold', 32)
        subtitle_font = load_font('bold', 24)
        text_font = load_font('regular', 20)

        # Верхний колонтитул
        header_height = 60
        draw.rectangle([(0, 0), (card_width, header_height)], fill=COLORS['header'])
        draw.text(
            (40, header_height // 2),
            "ТЕХНИЧЕСКАЯ КАРТОЧКА ТОВАРА",
            fill=COLORS['light_text'],
            font=title_font,
            anchor='lm'
        )

        # Основное содержимое
        content_start = header_height + 40
        y_pos = content_start

        # Изображение товара (уменьшенный размер)
        img = product_info['image'] or Image.open(PLACEHOLDER_IMAGE) if os.path.exists(PLACEHOLDER_IMAGE) else None

        if img:
            # Максимальные размеры для изображения
            max_img_width = 350
            max_img_height = 250

            # Масштабируем с сохранением пропорций
            img_ratio = img.width / img.height
            if img_ratio > 1:  # Горизонтальное
                new_width = min(max_img_width, img.width)
                new_height = int(new_width / img_ratio)
            else:  # Вертикальное
                new_height = min(max_img_height, img.height)
                new_width = int(new_height * img_ratio)

            img = img.resize((new_width, new_height), Image.LANCZOS)

            # Позиция изображения (слева)
            img_x = 40
            img_y = y_pos

            # Добавляем белую рамку
            bordered_img = ImageOps.expand(img, border=1, fill='white')
            # Вставляем изображение
            card.paste(bordered_img, (img_x, img_y))

            # Область для текста (справа от изображения)
            text_x = img_x + new_width + 40
            text_width = card_width - text_x - 40
        else:
            # Если нет изображения, текст занимает всю ширину
            text_x = 40
            text_width = card_width - 80

        # Заголовок артикула
        draw.text(
            (text_x, y_pos),
            f"Артикул: {article}",
            fill=COLORS['text'],
            font=subtitle_font
        )
        y_pos += 50

        # Разделительная линия
        draw.line([(text_x, y_pos), (text_x + text_width, y_pos)], fill=COLORS['header'], width=2)
        y_pos += 30

        # Заголовок характеристик
        draw.text(
            (text_x, y_pos),
            "ХАРАКТЕРИСТИКИ:",
            fill=COLORS['accent'],
            font=subtitle_font
        )
        y_pos += 40

        # Вывод характеристик в формате "Параметр: значение"
        for char in product_info['characteristics']:
            if ':' in char:
                param, value = char.split(':', 1)
                # Параметр (жирный)
                draw.text(
                    (text_x, y_pos),
                    f"{param.strip()}:",
                    fill=COLORS['text'],
                    font=load_font('bold', 20)
                )
                param_width = draw.textlength(f"{param.strip()}:", font=load_font('bold', 20))
                # Значение (обычный)
                draw.text(
                    (text_x + param_width + 10, y_pos),
                    value.strip(),
                    fill=COLORS['text'],
                    font=text_font
                )
                y_pos += 30

            if y_pos > card_height - 100:
                break

        # Нижний колонтитул
        footer_height = 40
        draw.rectangle(
            [(0, card_height - footer_height), (card_width, card_height)],
            fill=COLORS['footer']
        )
        draw.text(
            (card_width - 40, card_height - footer_height // 2),
            "https://agbgroup.ru",
            fill=COLORS['light_text'],
            font=load_font('regular', 18),
            anchor='rm'
        )

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
    print("=== Генератор технических карточек товаров ===")
    print("Поиск изображений через Yandex.eu | Четкий список характеристик")
    process_articles(INPUT_FILE)
    print("\n=== Обработка завершена ===")