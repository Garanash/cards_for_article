# -*- coding: utf-8 -*-
import os
import re
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

# Создаем папку для карточек
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# Инициализация клиента Perplexity
client = OpenAI(api_key=PERPLEXITY_API_KEY, base_url="https://api.perplexity.ai")


def sanitize_filename(filename):
    """Удаление недопустимых символов из имени файла"""
    return re.sub(r'[<>:"/\\|?*]', '_', filename)


def translit_to_ascii(text):
    """Транслитерация кириллицы в латиницу для поиска изображений"""
    translit_dict = {
        'а': 'a', 'б': 'b', 'в': 'v', 'г': 'g', 'д': 'd', 'е': 'e',
        'ё': 'yo', 'ж': 'zh', 'з': 'z', 'и': 'i', 'й': 'y', 'к': 'k',
        'л': 'l', 'м': 'm', 'н': 'n', 'о': 'o', 'п': 'p', 'р': 'r',
        'с': 's', 'т': 't', 'у': 'u', 'ф': 'f', 'х': 'h', 'ц': 'ts',
        'ч': 'ch', 'ш': 'sh', 'щ': 'sch', 'ъ': '', 'ы': 'y', 'ь': '',
        'э': 'e', 'ю': 'yu', 'я': 'ya'
    }

    result = []
    for char in text.lower():
        if char in translit_dict:
            result.append(translit_dict[char])
        elif char.isalnum():
            result.append(char)
        else:
            result.append(' ')

    return ' '.join(''.join(result).split())


def search_product_image(query):
    """Безопасный поиск изображения через Pexels API с обработкой Unicode"""
    try:
        # Транслитерируем запрос и кодируем для URL
        clean_query = translit_to_ascii(query)
        encoded_query = quote(clean_query)

        headers = {"Authorization": PEXELS_API_KEY}
        url = f"https://api.pexels.com/v1/search?query={encoded_query}&per_page=1"

        response = requests.get(url, headers=headers)
        response.raise_for_status()

        data = response.json()
        if data.get('photos') and len(data['photos']) > 0:
            return data['photos'][0]['src']['medium']
        return None
    except Exception as e:
        print(f"Ошибка поиска изображения: {str(e)}")
        return None


def perplexity_search(query):
    """Поиск через Perplexity API"""
    try:
        response = client.chat.completions.create(
            model="sonar-pro",
            messages=[
                {"role": "system", "content": "You are a helpful assistant."},
                {"role": "user", "content": query}
            ]
        )
        return response.choices[0].message.content
    except Exception as e:
        print(f"Ошибка поиска Perplexity: {e}")
        return None


def extract_product_info_from_perplexity(article):
    """Извлечение информации о товаре с помощью Perplexity API"""
    query = f"Найди полное описание и все характеристики артикула: {article}. Выведи только технические характеристики, без цены и других коммерческих предложений."
    search_results = perplexity_search(query)

    if not search_results:
        return None

    # Удаляем строки, содержащие слово "Цена" в любом регистре
    cleaned_details = []
    for line in search_results.split('\n'):
        if not re.search(r'цена', line, re.IGNORECASE):
            cleaned_details.append(line)
    cleaned_details = '\n'.join(cleaned_details)

    # Поиск изображения товара
    image_url = search_product_image(article)
    image_data = None
    if image_url:
        try:
            response = requests.get(image_url)
            image_data = Image.open(BytesIO(response.content))
        except Exception as e:
            print(f"Ошибка загрузки изображения: {e}")

    # Создаем словарь для хранения информации о товаре
    product_info = {
        'name': f"Артикул: {article}",
        'details': cleaned_details,
        'image': image_data,
        'source': 'Perplexity API + Pexels'
    }

    return product_info


def create_product_card(article, product_info):
    """Создание карточки товара"""
    try:
        # Размеры карточки
        card_width = 800
        card_height = 1200

        # Создаем белый фон
        card = Image.new('RGB', (card_width, card_height), 'white')
        draw = ImageDraw.Draw(card)

        # Загружаем шрифты
        try:
            title_font = ImageFont.truetype(FONT_PATH, 40)
            text_font = ImageFont.truetype(FONT_PATH, 28)
            small_font = ImageFont.truetype(FONT_PATH, 20)
        except:
            # Fallback на стандартные шрифты
            title_font = ImageFont.load_default(40)
            text_font = ImageFont.load_default(28)
            small_font = ImageFont.load_default(20)

        # Позиции для элементов
        y_position = 30

        # Добавляем изображение товара, если есть
        if product_info['image']:
            try:
                img = product_info['image']
                # Масштабируем изображение
                img_width, img_height = img.size
                new_width = card_width - 40
                new_height = int((new_width / img_width) * img_height)
                if new_height > 300:  # Ограничиваем максимальную высоту
                    new_height = 300
                    new_width = int((new_height / img_height) * img_width)

                img = img.resize((new_width, new_height), Image.LANCZOS)

                # Центрируем изображение
                x_position = (card_width - new_width) // 2
                card.paste(img, (x_position, y_position))
                y_position += new_height + 30
            except Exception as e:
                print(f"Ошибка обработки изображения: {e}")

        # Артикул (центрируем)
        text_width = draw.textlength(product_info['name'], font=title_font)
        x_position = (card_width - text_width) // 2
        draw.text((x_position, y_position), product_info['name'], fill='black', font=title_font)
        y_position += 60

        # Разделительная линия
        draw.line((20, y_position, card_width - 20, y_position), fill='gray', width=2)
        y_position += 30

        # Характеристики
        draw.text((30, y_position), "Характеристики:", fill='black', font=text_font)
        y_position += 50

        # Отображаем все характеристики
        for line in product_info['details'].split('\n'):
            if line.strip():  # Пропускаем пустые строки
                draw.text((40, y_position), line.strip(), fill='black', font=small_font)
                y_position += 30
                if y_position > card_height - 100:
                    break

        # Разделительная линия внизу
        draw.line((20, card_height - 60, card_width - 20, card_height - 60), fill='gray', width=1)

        # Источник
        source = product_info.get('source', 'Неизвестен')
        draw.text((30, card_height - 50), f"Источник: {source}", fill='gray', font=small_font)

        # Сохраняем
        safe_article = sanitize_filename(article)
        output_path = os.path.join(OUTPUT_FOLDER, f"{safe_article}.jpg")
        card.save(output_path, quality=90)
        print(f"Карточка сохранена: {output_path}")

        return True
    except Exception as e:
        print(f"Ошибка создания карточки: {e}")
        return False


def process_articles(file_path):
    """Обработка файла с артикулами"""
    try:
        df = pd.read_excel(file_path)
        articles = df.iloc[:, 1].dropna().astype(str).unique()

        print(f"Найдено {len(articles)} артикулов")

        for article in articles:
            article = article.strip()
            print(f"\nОбработка артикула: {article}")

            product_info = extract_product_info_from_perplexity(article)
            if product_info:
                create_product_card(article, product_info)
            else:
                print(f"Информация не найдена для {article}")

            time.sleep(DELAY)

    except Exception as e:
        print(f"Ошибка обработки файла: {e}")


if __name__ == '__main__':
    print("Запуск обработки артикулов...")
    process_articles(INPUT_FILE)
    print("Готово!")