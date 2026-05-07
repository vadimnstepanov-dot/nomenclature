#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт для выгрузки полного справочника номенклатуры из API
в Excel-файл со всеми доступными полями.

Дата: 2026-05-06
"""

import os
import sys
import time
import logging
import argparse
from datetime import datetime
from typing import Optional, Dict, List, Any
from pathlib import Path

import requests
import pandas as pd
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

try:
    from dotenv import load_dotenv

    load_dotenv()
except ImportError:
    pass

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('export.log', encoding='utf-8', mode='w'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)


class NomenclatureAPI:
    """Клиент для работы с API Эталонного номенклатура."""

    def __init__(
            self,
            base_url: str,
            token: Optional[str] = None,
            timeout: int = 30,
            max_retries: int = 3
    ):
        self.base_url = base_url.rstrip('/')
        self.timeout = timeout
        self.token = token  # Сохраняем токен для проверки
        self.session = self._create_session(max_retries)

        if token:
            self.session.headers.update({
                'Authorization': f'Bearer {token}',
                'Content-Type': 'application/json'
            })

    def _create_session(self, max_retries: int) -> requests.Session:
        """Создание сессии с настройками повторных запросов."""
        session = requests.Session()
        retry = Retry(
            total=max_retries,
            backoff_factor=0.3,
            status_forcelist=[429, 500, 502, 503, 504],
            allowed_methods=['GET', 'POST']
        )
        adapter = HTTPAdapter(max_retries=retry)
        session.mount('https://', adapter)
        session.mount('http://', adapter)
        return session

    def _request(self, endpoint: str, params: Optional[Dict] = None, require_auth: bool = False) -> Dict:
        """Выполнение GET-запроса к API."""
        # Проверка авторизации
        if require_auth and not self.token:
            raise PermissionError(f'Эндпоинт {endpoint} требует авторизации. Передайте токен.')

        url = f'{self.base_url}/{endpoint.lstrip("/")}'
        try:
            logger.debug(f'Запрос: {url} | Параметры: {params}')
            response = self.session.get(url, params=params, timeout=self.timeout)

            if response.status_code == 403:
                if not self.token:
                    raise PermissionError(
                        f'Доступ запрещён к {endpoint}. Требуется Bearer-токен. '
                        'Используйте параметр --token или переменную окружения API_TOKEN'
                    )
                else:
                    raise PermissionError(
                        f'Доступ запрещён к {endpoint}. Проверьте действительность токена.'
                    )

            response.raise_for_status()
            return response.json()
        except requests.exceptions.RequestException as e:
            logger.error(f'Ошибка запроса к {endpoint}: {e}')
            raise

    def get_categories(self) -> List[Dict]:
        """Получение всех категорий."""
        logger.info('Загрузка категорий...')
        all_categories = []
        page = 1
        limit = 100

        while True:
            data = self._request('/categories', {'page': page, 'limit': limit})
            items = data.get('data', [])
            all_categories.extend(items)

            meta = data.get('meta', {})
            total = meta.get('total', 0)

            logger.info(f'Загружено {len(all_categories)} из {total} категорий')

            if len(all_categories) >= total or not items:
                break
            page += 1
            time.sleep(0.2)

        return self._build_category_tree(all_categories)

    def _build_category_tree(self, categories: List[Dict]) -> Dict[int, Dict]:
        """Построение иерархии категорий."""
        cat_dict = {c['id']: {**c, 'path': c['name']} for c in categories}

        for cat in categories:
            parent = cat.get('category')
            if parent and isinstance(parent, dict):
                parent_id = parent.get('id')
                if parent_id and parent_id in cat_dict:
                    cat_dict[cat['id']]['path'] = f"{cat_dict[parent_id]['path']} > {cat['name']}"

        return cat_dict

    def get_units(self) -> Dict[str, Dict]:
        """Получение единиц измерения (требует авторизации)."""
        logger.info('Загрузка единиц измерения...')
        all_units = {}
        page = 1
        limit = 100

        while True:
            # require_auth=True для явной проверки авторизации
            data = self._request('/dictionary/synonym/units', {'page': page, 'limit': limit}, require_auth=True)
            for unit in data.get('data', []):
                all_units[unit['id']] = unit

            meta = data.get('meta', {})
            if len(all_units) >= meta.get('total', 0) or not data.get('data'):
                break
            page += 1
            time.sleep(0.2)

        return all_units

    def get_items_batch(self, page: int, limit: int, **filters) -> Dict:
        """Получение одной страницы эталонов."""
        params = {'page': page, 'limit': limit, **filters}
        params = {k: v for k, v in params.items() if v is not None}
        return self._request('/items', params)

    def get_all_items(
            self,
            batch_size: int = 100,
            status: Optional[str] = None,
            category_id: Optional[int] = None
    ) -> List[Dict]:
        """Получение всех эталонов с поддержкой пагинации."""
        logger.info(f'Начало выгрузки эталонов (партиями по {batch_size})...')
        all_items = []
        page = 1

        while True:
            try:
                data = self.get_items_batch(
                    page=page,
                    limit=batch_size,
                    status=status,
                    category=category_id
                )
                items = data.get('data', [])
                all_items.extend(items)

                meta = data.get('meta', {})
                total = meta.get('total', 0)

                logger.info(f'Страница {page}: загружено {len(all_items)} из {total} записей')

                if len(all_items) >= total or not items:
                    break
                page += 1
                time.sleep(0.3)

            except Exception as e:
                logger.error(f'Ошибка на странице {page}: {e}')
                if page > 1:
                    logger.warning('Продолжаем с предыдущей успешной страницы...')
                    break
                raise

        logger.info(f'Всего загружено эталонов: {len(all_items)}')
        return all_items


class DataTransformer:
    """Преобразование данных API в плоскую структуру для Excel."""

    @staticmethod
    def flatten_attribute_values(values: List[Dict]) -> Dict[str, Any]:
        """Преобразование списка значений атрибутов в словарь."""
        result = {}
        for val in values:
            attr = val.get('attribute', {})
            attr_name = attr.get('name', f"attr_{val.get('id')}")
            keyword = val.get('keyword', {})

            # Формируем ключ с префиксом для избежания конфликтов
            key = f"attr_{attr_name}"

            if keyword:
                result[key] = keyword.get('value')
            elif val.get('value'):
                result[key] = val.get('value')
            else:
                result[key] = None
        return result

    @staticmethod
    def flatten_standard(item: Dict, categories: Dict, units: Dict) -> Dict[str, Any]:
        """Преобразование одного эталона в плоскую структуру."""
        row = {}

        # Основные поля эталона
        row['id'] = item.get('id')
        row['name'] = item.get('name')
        row['comment'] = item.get('comment')
        row['active'] = item.get('active')
        row['confirmed'] = item.get('confirmed')
        row['viewed'] = item.get('viewed')

        # Даты
        for field in ['created', 'updated', 'deleted']:
            row[field] = item.get(field)

        # Категория
        category = item.get('category', {})
        row['category_id'] = category.get('id') if category else None
        row['category_name'] = category.get('name') if category else None
        row['category_path'] = categories.get(category.get('id'), {}).get('path') if categories and category else (
            category.get('name') if category else None)
        row['category_code'] = category.get('code') if category else None

        # Источник
        source = item.get('source', {})
        row['source_id'] = source.get('id') if source else None
        row['source_type'] = source.get('type') if source else None
        row['source_company'] = source.get('company') if source else None
        row['source_user'] = source.get('user') if source else None
        row['source_module'] = source.get('module') if source else None

        # Пользователи
        for role in ['creator', 'editor']:
            user = item.get(role, {})
            row[f'{role}_id'] = user.get('id') if user else None
            row[f'{role}_login'] = user.get('login') if user else None

        # Атрибуты (динамические поля)
        attr_values = DataTransformer.flatten_attribute_values(item.get('values', []))
        row.update(attr_values)

        # Изображения
        images = item.get('images', [])
        row['images_count'] = len(images) if images else 0
        row['image_paths'] = '; '.join([img.get('path', '') for img in images]) if images else None

        # Лог изменений
        log = item.get('log', {})
        row['log_description'] = log.get('description') if log else None
        row['log_type'] = log.get('type') if log else None

        # Связанные эталоны
        linked_item = item.get('item')
        if linked_item and isinstance(linked_item, dict):
            row['parent_standard_id'] = linked_item.get('id')
            row['parent_standard_name'] = linked_item.get('name')

        return row

    @staticmethod
    def to_dataframe(items: List[Dict], categories: Dict, units: Dict) -> pd.DataFrame:
        """Конвертация списка эталонов в DataFrame."""
        logger.info('Преобразование данных в таблицу...')
        flattened = [
            DataTransformer.flatten_standard(item, categories, units)
            for item in items
        ]
        df = pd.DataFrame(flattened)
        logger.info(f'Создан DataFrame: {df.shape[0]} строк × {df.shape[1]} колонок')
        return df


class ExcelExporter:
    """Экспорт данных в Excel с форматированием."""

    @staticmethod
    def export(
            df: pd.DataFrame,
            filepath: str,
            sheet_name: str = 'Номенклатура',
            freeze_first_row: bool = True
    ):
        """Экспорт DataFrame в Excel-файл."""
        logger.info(f'Экспорт в файл: {filepath}')

        # Создаем директорию если нужно
        Path(filepath).parent.mkdir(parents=True, exist_ok=True)

        # Экспорт с настройками
        with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)

            # Форматирование
            worksheet = writer.sheets[sheet_name]

            # Автоширина колонок (опционально, может замедлять на больших данных)
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)  # Ограничение ширины
                worksheet.column_dimensions[column_letter].width = adjusted_width

            # Закрепление первой строки
            if freeze_first_row:
                worksheet.freeze_panes = 'A2'

        logger.info(f'Файл успешно сохранён: {filepath}')


def main():
    """Точка входа в скрипт."""
    parser = argparse.ArgumentParser(
        description='Выгрузка справочника номенклатуры в Excel',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
Примеры использования:
  %(prog)s --output my_export.xlsx
  %(prog)s --status active --category 42 --batch-size 50
  %(prog)s --token YOUR_TOKEN --output ./data/export_$(date).xlsx
        '''
    )

    parser.add_argument(
        '--output', '-o',
        default=os.getenv('OUTPUT_FILE', 'nomenclature_export.xlsx'),
        help='Путь к выходному Excel-файлу'
    )
    parser.add_argument(
        '--url', '-u',
        default=os.getenv('API_BASE_URL', 'https://zakupki.rzd-medicine.ru/nomenklator/api'),
        help='Базовый URL API'
    )
    parser.add_argument(
        '--token', '-t',
        default=os.getenv('API_TOKEN'),
        help='Bearer-токен авторизации'
    )
    parser.add_argument(
        '--status', '-s',
        choices=['active', 'inactive', 'confirmed'],
        help='Фильтр по статусу эталонов'
    )
    parser.add_argument(
        '--category', '-c',
        type=int,
        help='ID категории для фильтрации'
    )
    parser.add_argument(
        '--batch-size', '-b',
        type=int,
        default=int(os.getenv('BATCH_SIZE', 100)),
        help='Размер партии записей (по умолчанию: 100)'
    )
    parser.add_argument(
        '--max-retries', '-r',
        type=int,
        default=int(os.getenv('MAX_RETRIES', 3)),
        help='Максимальное количество повторных попыток запроса'
    )
    parser.add_argument(
        '--dry-run', '-d',
        action='store_true',
        help='Тестовый режим: загрузить только первые 10 записей'
    )

    args = parser.parse_args()

    # Инициализация
    start_time = time.time()
    logger.info('=== Запуск выгрузки номенклатуры ===')

    try:
        # Создание клиента API
        api = NomenclatureAPI(
            base_url=args.url,
            token=args.token,
            max_retries=args.max_retries
        )

        # 1. Загрузка категорий (не требует авторизации)
        categories = api.get_categories()

        # 2. Загрузка единиц измерения (требует авторизации - опционально)
        units = {}
        if args.token:
            try:
                units = api.get_units()
                logger.info(f'✅ Загружено {len(units)} единиц измерения')
            except PermissionError as e:
                logger.warning(f'⚠ Не удалось загрузить единицы измерения: {e}')
                logger.warning('💡 Данные будут экспортированы без справочника единиц')
            except Exception as e:
                logger.warning(f'⚠ Ошибка при загрузке единиц измерения: {e}')
                logger.warning('💡 Продолжаем выгрузку без единиц измерения')
        else:
            logger.info('ℹ Токен не указан — пропускаем загрузку единиц измерения')

        # 3. Загрузка эталонов
        items = api.get_all_items(
            batch_size=10 if args.dry_run else args.batch_size,
            status=args.status,
            category_id=args.category
        )

        if not items:
            logger.warning('⚠ Нет данных для экспорта')
            return

        # 4. Преобразование данных
        df = DataTransformer.to_dataframe(items, categories, units)

        # 5. Экспорт в Excel
        ExcelExporter.export(df, args.output)

        # 6. Статистика
        elapsed = time.time() - start_time
        units_status = "✅ загружены" if units else "⚠ пропущены (требуется токен)"

        logger.info(f'''
=== Выгрузка завершена ===
✅ Записей: {len(df)}
✅ Колонок: {len(df.columns)}
⏱ Время выполнения: {elapsed:.1f} сек.
📁 Файл: {Path(args.output).resolve()}
🔑 Единицы измерения: {units_status}
        ''')

    except PermissionError as e:
        logger.error(f'❌ Ошибка авторизации: {e}')
        logger.error('💡 Решение: передайте корректный токен через --token или API_TOKEN')
        sys.exit(403)
    except KeyboardInterrupt:
        logger.warning('⚠ Прервано пользователем')
        sys.exit(130)
    except Exception as e:
        logger.exception(f'❌ Критическая ошибка: {e}')
        sys.exit(1)


if __name__ == '__main__':
    main()
