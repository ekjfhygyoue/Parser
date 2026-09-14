from __future__ import annotations
import requests
from bs4 import BeautifulSoup
import os
import pandas
from io import BytesIO
from sqlalchemy import create_engine
import xlrd
import re
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path


class html_converter:

    def __init__(self, data_start, data_end, link):
        self.data_start = data_start
        self.data_end = data_end
        self.link = link

    # Метод загружает и парсит HTML-страницу по указанному URL с параметром пагинации.
    def get_page(self, number_page):
        parametr = {
            'page': f'page-{number_page}'
        }
        result = self.open_link(self.link, parametr)
        if result != 0:
            soup = BeautifulSoup(result.text, 'html.parser')
            print(f'Страница под номером {number_page} обработана')
            return soup
        else:
            return None

    # Метод извлекает все ссылки на XLS-файлы из HTML-страницы, используя BeautifulSoup.
    def get_links_with_file(self, soup: BeautifulSoup):
        elements = soup.find_all('div', class_='accordeon-inner__item')

        results = []

        for item in elements:
            link_tag = item.find('a', class_='xlsx')
            big_data_tag = item.find('div', class_='accordeon-inner__item-inner__title')
            date_tag = None
            if big_data_tag:
                date_tag = big_data_tag.find('p')


            if link_tag and date_tag:
                xlsx_href = link_tag.get('href')

                if xlsx_href and not xlsx_href.startswith('http'):
                    xlsx_url = 'https://spimex.com' + xlsx_href
                else:
                    xlsx_url = xlsx_href

                trade_date = date_tag.text.strip()

                results.append({
                    'date': trade_date,
                    'xlsx_link': xlsx_url
                })
        return results

    def download_all_xlsx(self):
        pages = self.get_page(1)
        links = self.get_links_with_file(pages)
        for item in links:
            link = item['xlsx_link']
            year = int(item['date'].split(' ')[1])
            if year >= self.data_start and year <= self.data_end:
                downloaded = self.read_excel_file(link)
                return downloaded


    def open_link(self, link, parametr = None):

        try:
            response = requests.get(link, params=parametr, timeout=300)
        except requests.exceptions.RequestException as e:
            print(f'Ошибка сети при запросе {e}')
            return 0
        if response.status_code != 200:
            print(f'Ошибка API (Код {response.status_code}). Ответ сервера:\n{response.text[:200]}')
            return 0
        return response





    def read_sheet_rows(self, url):
        if isinstance(url, str) and url.startswith('http'):
            self.open_link(url)
        else:
            with open(url, 'rb') as f:
                data = f.read()


    def detect_engine(data: bytes) -> str:
        if data.startswith(b'\xd0\xcf\x11\xe0'):
            return 'xlrd'  # настоящий .xls
        if data.startswith(b'PK\x03\x04'):
            return 'openpyxl'  # .xlsx
        if data.startswith(b'%PDF'):
            raise ValueError('Это PDF, а не Excel')
        if data.startswith(b'<html') or data.startswith(b'<!DOC'):
            raise ValueError('Сайт отдал HTML вместо файла')
        raise ValueError(f'Неизвестный формат: {data[:8].hex(" ")}')

    def read_excel_file(self, link):
        response = requests.get(link, timeout=60)
        if response.status_code != 200:
            print(f'Не скачалось: {link}, код {response.status_code}')
            return None

        content = response.content
        engine = self.detect_engine(content)  # ← bytes, не BytesIO
        print(f'Движок: {engine}')

        xl_file = pandas.ExcelFile(BytesIO(content), engine=engine)
        print(f'Листы: {xl_file.sheet_names}')
        print('=' * 60)

        all_sheets = {}
        for sheet_name in xl_file.sheet_names:
            print(f"\n📄 Лист: '{sheet_name}'")

            df_raw = pandas.read_excel(
                BytesIO(content),  # ← новый BytesIO
                sheet_name=sheet_name,
                header=None,
                engine=engine,  # ← тот же движок
            )
            all_sheets[sheet_name] = df_raw

            print(f"Первые 20 строк листа '{sheet_name}':")
            for i in range(min(20, len(df_raw))):
                row = df_raw.iloc[i]
                row_text = ' | '.join(
                    str(cell) for cell in row if pandas.notna(cell)
                )
                if row_text.strip():
                    print(f"  Строка {i}: {row_text[:200]}")
            print('-' * 60)

        return all_sheets

if __name__=='__main__':
    test = html_converter(2023, 2026, 'https://spimex.com/markets/oil_products/trades/results/')
    test.download_all_xlsx()