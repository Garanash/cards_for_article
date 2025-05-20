import os
import re
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
from io import BytesIO
import time
from openai import OpenAI

# Настройки
INPUT_FILE = 'articles.xlsx'  # Файл с артикулами (первый лист, второй столбец)
OUTPUT_FOLDER = 'product_cards'  # Папка для сохранения карточек
FONT_PATH = 'arial.ttf'  # Путь к шрифту
PERPLEXITY_API_KEY = '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'  # Ваш API ключ Perplexity
DELAY = 2  # Задержка между запросами (секунды)

# Создаем папку для карточек
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# Инициализация клиента Perplexity
client = OpenAI(api_key=PERPLEXITY_API_KEY, base_url="https://api.perplexity.ai")

def sanitize_filename(filename):
    """Удаление недопустимых символов из имени файла"""
    return re.sub(r'[<>:"/\\|?*]', '_', filename)

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
    query = f"найди все характеристики артикула: {article} и выведи их через двоеточие, без дополнительных комментариев"
    search_results = perplexity_search(query)

    if not search_results:
        return None

    # Создаем словарь для хранения информации о товаре
    product_info = {
        'name': f"Артикул: {article}",
        'price': 'Цена не указана',
        'details': search_results,  # Сохраняем весь ответ как есть
        'image': None,
        'source': 'Perplexity API'
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
            title_font = ImageFont.truetype(FONT_PATH, 36)
            text_font = ImageFont.truetype(FONT_PATH, 24)
            small_font = ImageFont.truetype(FONT_PATH, 18)
        except:
            title_font = ImageFont.load_default()
            text_font = ImageFont.load_default()
            small_font = ImageFont.load_default()

        # Позиции для элементов
        y_position = 20

        # Артикул
        draw.text((20, y_position), product_info['name'], fill='black', font=text_font)
        y_position += 40

        # Цена
        draw.text((20, y_position), f"Цена: {product_info['price']}", fill='red', font=text_font)
        y_position += 50

        # Характеристики
        draw.text((20, y_position), "Характеристики:", fill='black', font=text_font)
        y_position += 40

        # Отображаем все характеристики как есть
        for line in product_info['details'].split('\n'):
            draw.text((30, y_position), line, fill='black', font=small_font)
            y_position += 30
            if y_position > card_height - 50:
                break

        # Источник
        source = product_info.get('source', 'Неизвестен')
        draw.text((20, card_height - 40), f"Источник: {source}", fill='gray', font=small_font)

        # Сохраняем
        safe_article = sanitize_filename(article)
        output_path = os.path.join(OUTPUT_FOLDER, f"{safe_article}.jpg")
        card.save(output_path)
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
