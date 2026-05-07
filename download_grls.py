#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скачивание справочника ГРЛС с grls.rosminzdrav.ru
Прямая загрузка ZIP-архива с XLS файлами разных статусов
"""

import requests
import pandas as pd
import logging
import zipfile
import io
import re
from pathlib import Path
from datetime import datetime

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


# Возможные статусы ГРЛС
VALID_STATUSES = [
    "Выдано по правилам ЕАЭС",
    "Действует, в иностранных упаковках",
    "Действует, на подтверждении государственной регистрации",
    "Действующий",
    "Изменённый",
    "Исключённый",
    "Истёкший",
    "Приостановлено применение"
]


class GRLSFullDownloader:
    """Скачивание справочника ГРЛС одним запросом (ZIP с XLS файлами)"""

    # Прямая ссылка на файл экспорта ГРЛС
    DOWNLOAD_URL = "https://grls.rosminzdrav.ru/GetGRLS.ashx?FileGUID=839375f7-3a90-4dcc-af8c-fe324ca0f9eb&UserReq=4281387"

    HEADERS = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
        'Accept': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/octet-stream, */*',
        'Accept-Language': 'ru-RU,ru;q=0.9',
        'Connection': 'keep-alive',
    }

    def __init__(self, output_dir: str = 'grls_export'):
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.session = requests.Session()
        self.session.headers.update(self.HEADERS)
        # Увеличиваем таймаут для больших данных
        self.timeout = 600

    def normalize_status(self, status_raw: str) -> str:
        """Нормализация статуса к одному из допустимых значений"""
        if not status_raw:
            return "Неизвестно"
        
        status_raw = status_raw.strip()
        
        # Точное совпадение
        if status_raw in VALID_STATUSES:
            return status_raw
        
        # Поиск частичного совпадения
        status_lower = status_raw.lower()
        
        for valid_status in VALID_STATUSES:
            if valid_status.lower() in status_lower or status_lower in valid_status.lower():
                return valid_status
        
        # Проверка по ключевым словам
        if any(word in status_lower for word in ['еаэс', 'правилам еаэс']):
            return "Выдано по правилам ЕАЭС"
        elif any(word in status_lower for word in ['иностранн', 'упаковк']):
            return "Действует, в иностранных упаковках"
        elif any(word in status_lower for word in ['подтвержд', 'регистрац']):
            return "Действует, на подтверждении государственной регистрации"
        elif 'действующ' in status_lower or 'действует' in status_lower:
            return "Действующий"
        elif 'изменён' in status_lower or 'изменен' in status_lower:
            return "Изменённый"
        elif 'исключ' in status_lower:
            return "Исключённый"
        elif 'истёк' in status_lower or 'истек' in status_lower or 'истекш' in status_lower:
            return "Истёкший"
        elif 'приостанов' in status_lower:
            return "Приостановлено применение"
        
        return "Неизвестно"

    def extract_status_from_sheet(self, df: pd.DataFrame) -> str:
        """Извлечение статуса из файла с учетом объединенных ячеек"""
        if len(df) == 0:
            return "Неизвестно"
        
        # Статус находится в строке 5 (индекс 5), колонка 2 или nearby
        # Проверяем строку 5 и соседние
        rows_to_check = [5, 4, 6, 3, 7]
        
        for row_idx in rows_to_check:
            if row_idx >= len(df):
                continue
            row = df.iloc[row_idx]
            for val in row.values:
                if isinstance(val, str) and val.strip():
                    normalized = self.normalize_status(val.strip())
                    if normalized != "Неизвестно":
                        logger.debug(f"   → Статус найден в строке {row_idx + 1}: {val.strip()}")
                        return normalized
        
        # Если не нашли в конкретных строках, проверяем все первые 10 строк
        rows_to_check = min(len(df), 10)
        for row_idx in range(rows_to_check):
            row = df.iloc[row_idx]
            for val in row.values:
                if isinstance(val, str) and val.strip():
                    normalized = self.normalize_status(val.strip())
                    if normalized != "Неизвестно":
                        logger.debug(f"   → Статус найден в строке {row_idx + 1}: {val.strip()}")
                        return normalized
        
        return "Неизвестно"

    def extract_upload_date_from_last_row(self, df: pd.DataFrame) -> str | None:
        """Извлечение даты выгрузки из последней строки файла"""
        if len(df) == 0:
            return None
        
        last_row = df.iloc[-1]
        
        for val in last_row.values:
            if isinstance(val, str):
                # Ищем дату в формате ДД.ММ.ГГГГ или ДД.ММ.ГГ
                date_pattern = r'(\d{1,2}\.\d{1,2}\.\d{2,4})'
                match = re.search(date_pattern, val)
                if match:
                    return match.group(1)
                
                # Проверяем, не является ли вся строка датой
                try:
                    # Попытка распознать дату в различных форматах
                    if re.match(r'^\d{1,2}\.\d{1,2}\.\d{2,4}$', val.strip()):
                        return val.strip()
                except:
                    pass
            
            # Проверяем, не является ли значение датой pandas
            elif isinstance(val, (pd.Timestamp, datetime)):
                return val.strftime('%d.%m.%Y')
        
        return None

    def process_xls_file(self, zf, xls_file: str) -> pd.DataFrame | None:
        """Обработка одного XLS файла из архива"""
        logger.info(f"📄 Обработка файла: {xls_file}")
        
        try:
            # Читаем файл правильно: заголовок из 5-й строки (index=4)
            if xls_file.endswith('.xlsx'):
                df = pd.read_excel(zf.open(xls_file), engine='openpyxl', header=4)
            else:
                df = pd.read_excel(zf.open(xls_file), engine='xlrd', header=4)
            
            if len(df) == 0:
                logger.warning(f"⚠️ Файл {xls_file} пустой после чтения с заголовком, пропускаем")
                return None
            
            # Читаем первые строки файла без заголовка для определения статуса
            # Статус обычно находится в строках 2-6 (до заголовка)
            if xls_file.endswith('.xlsx'):
                df_raw = pd.read_excel(zf.open(xls_file), engine='openpyxl', header=None, nrows=10)
            else:
                df_raw = pd.read_excel(zf.open(xls_file), engine='xlrd', header=None, nrows=10)
            
            # Определяем статус из первых строк
            status = self.extract_status_from_sheet(df_raw)
            logger.info(f"   → Статус: {status}")
            
            # Извлекаем дату выгрузки из последних строк исходного файла
            # Для этого читаем весь файл еще раз без заголовка
            if xls_file.endswith('.xlsx'):
                df_full_raw = pd.read_excel(zf.open(xls_file), engine='openpyxl', header=None)
            else:
                df_full_raw = pd.read_excel(zf.open(xls_file), engine='xlrd', header=None)
            
            upload_date = self.extract_upload_date_from_last_row(df_full_raw)
            
            # Удаляем последние строки, если они содержат только дату выгрузки
            # Проверяем последние 2 строки
            rows_to_drop = []
            for i in range(len(df) - 1, max(len(df) - 3, -1), -1):
                row = df.iloc[i]
                # Если строка содержит только одну непустую ячейку с датой или похожа на служебную
                non_empty = sum(1 for v in row.values if isinstance(v, str) and v.strip())
                if non_empty <= 2:
                    # Проверяем, есть ли дата в этой строке
                    has_date = False
                    for v in row.values:
                        if isinstance(v, str) and re.search(r'\d{1,2}\.\d{1,2}\.\d{2,4}', v):
                            has_date = True
                            break
                    if has_date:
                        rows_to_drop.append(i)
            
            if rows_to_drop:
                df = df.drop(rows_to_drop).reset_index(drop=True)
            
            # Удаляем первые 2 столбца (лишние, Unnamed)
            if len(df.columns) > 2:
                df = df.iloc[:, 2:]
            
            # Проверяем первую строку - если она содержит значение статуса, удаляем её
            if len(df) > 0:
                first_row_values = df.iloc[0].values
                for val in first_row_values:
                    if isinstance(val, str) and val.strip() in VALID_STATUSES:
                        logger.info("   → Удалена строка со статусом из данных")
                        df = df.drop(0).reset_index(drop=True)
                        break
            
            # Добавляем столбец со статусом
            df['Статус'] = status
            
            # Добавляем столбец с датой выгрузки
            df['Дата выгрузки'] = upload_date if upload_date else None
            
            return df
            
        except Exception as e:
            logger.error(f"❌ Ошибка обработки файла {xls_file}: {e}")
            return None

    def close(self):
        self.session.close()


    def download_and_process(self, timeout: int = 600) -> tuple[Path | None, pd.DataFrame | None]:
        """Скачивание ZIP и обработка всех XLS файлов. Возвращает (путь к ZIP, объединённый DataFrame)"""
        logger.info("📥 Загрузка данных из ГРЛС...")

        try:
            response = self.session.get(
                self.DOWNLOAD_URL,
                timeout=timeout
            )
            response.raise_for_status()

            content_type = response.headers.get('Content-Type', '').lower()
            logger.info(f"Content-Type от сервера: {content_type} (может быть неверным)")
            
            # Проверяем, что это ZIP по магическим байтам
            magic = response.content[:4]
            if magic != b'\x50\x4B\x03\x04':
                logger.warning(f"⚠️ Файл не является ZIP. Первые 4 байта (hex): {magic.hex()}")
                # Пробуем продолжить, возможно это всё равно ZIP
            
            # Сохраняем ZIP файл
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            zip_path = self.output_dir / f'grls_full_{timestamp}.zip'

            with open(zip_path, 'wb') as f:
                f.write(response.content)

            file_size = zip_path.stat().st_size / 1024 / 1024  # MB
            logger.info(f"✅ ZIP сохранён: {zip_path} ({file_size:.2f} MB)")

            # Распаковываем и обрабатываем все XLS файлы
            all_dfs = []
            
            with zipfile.ZipFile(zip_path, 'r') as zf:
                xls_files = [f for f in zf.namelist() if f.endswith(('.xls', '.xlsx'))]
                
                if not xls_files:
                    logger.error("❌ В ZIP архиве не найдено XLS/XLSX файлов")
                    return zip_path, None
                
                logger.info(f"📂 Найдено файлов в архиве: {len(xls_files)}")
                
                for xls_file in xls_files:
                    df = self.process_xls_file(zf, xls_file)
                    if df is not None:
                        all_dfs.append(df)

            if not all_dfs:
                logger.error("❌ Не удалось прочитать ни один файл")
                return zip_path, None

            # Объединяем все DataFrame
            combined_df = pd.concat(all_dfs, ignore_index=True)
            logger.info(f"✅ Объединено записей: {len(combined_df)}")
            
            return zip_path, combined_df

        except requests.exceptions.Timeout:
            logger.error("❌ Таймаут запроса. Попробуйте увеличить timeout")
            return None, None
        except requests.exceptions.RequestException as e:
            logger.error(f"❌ Ошибка запроса: {e}")
            return None, None
        except zipfile.BadZipFile as e:
            logger.error(f"❌ Ошибка ZIP архива: {e}")
            return None, None
        except Exception as e:
            logger.error(f"❌ Ошибка обработки: {e}")
            return None, None


def main():
    """Основная функция"""
    logger.info("🚀 Запуск скачивания справочника ГРЛС")

    downloader = GRLSFullDownloader()

    try:
        # Скачивание и обработка ZIP с XLS файлами
        zip_path, df = downloader.download_and_process(timeout=600)
        
        if df is not None and len(df) > 0:
            print(f"\n📊 Статистика:")
            print(f"   • Записей: {len(df)}")
            print(f"   • Колонок: {len(df.columns)}")
            print(f"   • Пример колонок: {list(df.columns)[:15]}")
            
            # Статистика по статусам
            if 'Статус' in df.columns:
                print(f"\n📋 Статусы записей:")
                status_counts = df['Статус'].value_counts()
                for status, count in status_counts.items():
                    print(f"   • {status}: {count}")

            # Сохранение в output/esklp/esklp_full.xlsx на лист ГРЛС
            output_dir = Path('output/esklp')
            output_dir.mkdir(parents=True, exist_ok=True)
            esklp_full_path = output_dir / 'esklp_full.xlsx'
            
            logger.info(f"📝 Сохранение данных на лист 'ГРЛС' в {esklp_full_path}")
            
            # Если файл существует, читаем его и добавляем/обновляем лист
            if esklp_full_path.exists():
                with pd.ExcelWriter(esklp_full_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                    df.to_excel(writer, sheet_name='ГРЛС', index=False)
            else:
                # Создаем новый файл с листом ГРЛС
                with pd.ExcelWriter(esklp_full_path, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name='ГРЛС', index=False)
            
            logger.info(f"✅ Данные сохранены в {esklp_full_path} на лист 'ГРЛС'")

    except KeyboardInterrupt:
        logger.warning("⚠️ Прервано пользователем")
    finally:
        downloader.close()
        logger.info("✅ Готово")


if __name__ == '__main__':
    main()