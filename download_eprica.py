#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт загрузки данных из шины ePrica.

Особенности работы с API ePrica:
- Данные передаются порциями не более 1000 позиций за раз
- Используется параметр RV (Record Version) для пагинации
- RV формируется как TimeStamp(UTC) + случайное число
- При первом подключении RV = 0
- Цикл повторяется до тех пор, пока запрос возвращает хотя бы одну строку
- Для полной перезагрузки RV устанавливается в 0

Автор: Assistant
Дата: 2026
"""

import requests
import pandas as pd
import json
import sys
import logging
import time
from collections import OrderedDict
from typing import Optional, List, Dict, Any, Tuple
from datetime import datetime, timezone

# ============================================================================
# КОНФИГУРАЦИЯ
# ============================================================================

AUTH_URL = "https://api.f3bus.ru/User/auth/"
GOODS_BASE = "https://api.f3bus.ru/esadditional/goods/"
LOGIN = "demo"
PASSWORD = "demo"
MAX_PAGE_SIZE = 1000  # Максимальный размер порции данных
MAX_RETRIES = 3  # Максимальное количество попыток при ошибке
RETRY_DELAY = 5  # Задержка между попытками в секундах
REQUEST_TIMEOUT = 60  # Таймаут запроса в секундах

# Точный список 85 полей из документации
EXPECTED_COLUMNS = [
    "codeEs", "guidEs", "guidEsNew", "Name", "trnNameRus", "trnNameLat",
    "groupName", "brandName", "cureformNameFull", "cureformNameLat",
    "cureformCode", "releaseForm", "weight", "weightUnit", "dosage",
    "characteristicsName", "packing", "regNumber", "regDate", "regDateEnd",
    "barcode", "regOwnerProducer", "regOwnerCountry", "producer",
    "producerLat", "country", "country_prod", "mnnRus", "mnnLat", "ndsRate",
    "jnvls", "pkkn", "notRecept", "isAlcoholContent", "isImperativeAssortiment",
    "isImmunobiological", "registerPrice", "mkB10", "guids", "ids",
    "attributes", "discribe", "codeOkpd2", "codeTnVed", "atcCode", "atcName",
    "storageCondition", "storageConditionTempMin", "storageConditionTempMax",
    "transportationTempMin", "transportationTempMax", "transportationHumidity",
    "Articul", "expirationDate", "packageLength", "packageHeight",
    "packageDeth", "packVolume", "packWeight", "boxPackQuantity", "gtin",
    "esklp", "seasone", "storeHumidity", "isLs", "isBad", "composition",
    "insertDate", "deleteDate", "isExplosive", "isFlammable", "isNarcotic",
    "isPsychotropicDrugs", "isPowerfull", "isHerbalMedicine",
    "volumeWeightNumberOfDoes", "instructionId", "producerLekform",
    "addressOwner", "ktru", "allGroup", "atcHierarchy", "grlsGroup",
    "grlsGuid"
]

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-8s | %(message)s",
    datefmt="%H:%M:%S"
)
logger = logging.getLogger(__name__)


# ============================================================================
# ФУНКЦИИ АВТОРИЗАЦИИ
# ============================================================================

def get_token() -> str:
    """
    Получение токена авторизации
    
    Returns:
        str: Токен доступа
        
    Raises:
        SystemExit: Если авторизация не удалась
    """
    try:
        logger.info("🔐 Авторизация в системе ePrica...")
        r = requests.post(
            AUTH_URL, 
            json={"login": LOGIN, "password": PASSWORD}, 
            timeout=REQUEST_TIMEOUT
        )
        r.raise_for_status()
        data = r.json()
        
        # Поиск токена в различных форматах ответа
        token = (
            data.get("token") or 
            data.get("access_token") or 
            (list(data.values())[0] if isinstance(data, dict) else None)
        )
        
        if not token:
            raise ValueError("Токен не найден в ответе сервера")
            
        logger.info("✅ Авторизация успешна")
        return token
        
    except requests.exceptions.RequestException as e:
        logger.error(f"❌ Ошибка сети при авторизации: {e}")
        sys.exit(1)
    except Exception as e:
        logger.error(f"❌ Ошибка авторизации: {type(e).__name__}: {e}")
        sys.exit(1)


# ============================================================================
# ФУНКЦИИ ЗАПРОСА ДАННЫХ
# ============================================================================

def fetch_page(token: str, rv: int = 0) -> Tuple[List[Dict[str, Any]], int]:
    """
    Запрос одной порции данных с указанным RV
    
    Args:
        token: Токен авторизации
        rv: Версия последней полученной записи (0 для первого запроса)
        
    Returns:
        Tuple[List[Dict], int]: Кортеж из списка данных и нового значения RV
    """
    headers = {
        "Authorization": f"Bearer {token}", 
        "Accept": "application/json"
    }
    
    url = f"{GOODS_BASE}{rv}"
    
    for attempt in range(MAX_RETRIES):
        try:
            logger.debug(f"📥 Запрос к API: {url} (попытка {attempt + 1}/{MAX_RETRIES})")
            
            r = requests.get(
                url, 
                headers=headers, 
                timeout=REQUEST_TIMEOUT
            )
            r.raise_for_status()
            
            data = r.json()
            
            # Распаковка возможных оберток ответа
            if isinstance(data, dict):
                # Извлекаем данные из обёртки
                items = None
                new_rv = rv
                
                for k in ["data", "items", "result", "goods"]:
                    if isinstance(data.get(k), list):
                        items = data[k]
                        break
                
                # Получаем новое значение RV из ответа
                if "rv" in data:
                    new_rv = data["rv"]
                elif "RV" in data:
                    new_rv = data["RV"]
                elif "nextRv" in data:
                    new_rv = data["nextRv"]
                    
                if items is None:
                    items = []
                    
                return items, new_rv
                
            elif isinstance(data, list):
                return data, rv
            else:
                logger.warning(f"⚠️ Неожиданный формат ответа: {type(data)}")
                return [], rv
                
        except requests.exceptions.Timeout:
            logger.warning(f"⚠️ Таймаут запроса (попытка {attempt + 1}/{MAX_RETRIES})")
            if attempt < MAX_RETRIES - 1:
                time.sleep(RETRY_DELAY * (attempt + 1))
            continue
            
        except requests.exceptions.RequestException as e:
            logger.warning(f"⚠️ Ошибка сети (попытка {attempt + 1}/{MAX_RETRIES}): {e}")
            if attempt < MAX_RETRIES - 1:
                time.sleep(RETRY_DELAY * (attempt + 1))
            continue
            
        except json.JSONDecodeError as e:
            logger.error(f"❌ Ошибка парсинга JSON: {e}")
            return [], rv
            
        except Exception as e:
            logger.warning(f"⚠️ Неожиданная ошибка (попытка {attempt + 1}/{MAX_RETRIES}): {type(e).__name__}: {e}")
            if attempt < MAX_RETRIES - 1:
                time.sleep(RETRY_DELAY * (attempt + 1))
            continue
    
    # Если все попытки исчерпаны
    logger.error(f"❌ Не удалось получить данные после {MAX_RETRIES} попыток")
    return [], rv


def fetch_all_data(token: str) -> List[Dict[str, Any]]:
    """
    Полная выгрузка всех данных с использованием пагинации через RV
    
    Args:
        token: Токен авторизации
        
    Returns:
        List[Dict]: Список всех записей
    """
    all_data = []
    current_rv = 0
    page_count = 0
    total_records = 0
    
    logger.info("🚀 Начало полной выгрузки данных из ePrica...")
    logger.info(f"📊 Параметр RV初始: {current_rv}")
    logger.info(f"📦 Максимальный размер порции: {MAX_PAGE_SIZE}")
    
    start_time = datetime.now()
    
    while True:
        # Запрос очередной порции данных
        page_data, new_rv = fetch_page(token, current_rv)
        
        page_count += 1
        records_in_page = len(page_data)
        total_records += records_in_page
        
        logger.info(
            f"📄 Страница {page_count}: получено {records_in_page} записей "
            f"(всего: {total_records}, RV: {current_rv} → {new_rv})"
        )
        
        # Если данных нет - завершаем цикл
        if not page_data:
            logger.info("🏁 Получены все данные (пустая страница)")
            break
        
        # Добавляем данные в общий список
        all_data.extend(page_data)
        
        # Обновляем RV для следующего запроса
        current_rv = new_rv
        
        # Проверка на зацикливание (если RV не изменился)
        if current_rv == 0 and page_count > 1:
            logger.warning("⚠️ RV вернулся в 0 после первой страницы - возможна ошибка API")
            break
            
        # Небольшая пауза между запросами для снижения нагрузки на API
        time.sleep(0.5)
    
    elapsed = (datetime.now() - start_time).total_seconds()
    
    logger.info("=" * 60)
    logger.info(f"✅ Выгрузка завершена")
    logger.info(f"📊 Всего страниц: {page_count}")
    logger.info(f"📊 Всего записей: {total_records}")
    logger.info(f"⏱️  Время выполнения: {elapsed:.1f} сек")
    logger.info(f"📈 Средняя скорость: {total_records / elapsed:.1f} записей/сек")
    logger.info("=" * 60)
    
    return all_data


# ============================================================================
# ФУНКЦИИ ОБРАБОТКИ И НОРМАЛИЗАЦИИ ДАННЫХ
# ============================================================================

def normalize_item(item: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    """
    Нормализация одной записи: приведение ключей к ожидаемым именам
    
    Args:
        item: Исходная запись
        
    Returns:
        Dict или None если запись некорректна
    """
    if not isinstance(item, dict):
        return None
    
    # Получаем GUID для идентификации записи
    guid = item.get("guidEs") or item.get("guides") or ""
    if not guid:
        return None
    
    # Карта соответствия: нормализованный ключ -> целевое имя колонки
    col_map = {col.lower().replace('с', 'c').strip(): col for col in EXPECTED_COLUMNS}
    
    row = {}
    for k, v in item.items():
        # Нормализация ключей (кириллица -> латиница, lower)
        norm_k = k.lower().replace('с', 'c').strip()
        target_key = col_map.get(norm_k, k)
        
        # Сериализация вложенных структур для Excel
        if isinstance(v, (dict, list)):
            row[target_key] = json.dumps(v, ensure_ascii=False)
        else:
            row[target_key] = v
    
    return row


def normalize_and_deduplicate(data: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Нормализация и дедупликация данных по guidEs
    
    Args:
        data: Список исходных записей
        
    Returns:
        List[Dict]: Список нормализованных уникальных записей
    """
    merged = OrderedDict()
    
    for item in data:
        normalized = normalize_item(item)
        if normalized:
            guid = normalized.get("guidEs", "")
            # Первая запись с таким GUID сохраняется (keep first)
            if guid not in merged:
                merged[guid] = normalized
    
    return list(merged.values())


# ============================================================================
# ФУНКЦИИ ЭКСПОРТА В EXCEL
# ============================================================================

def export_to_dataframe(data: List[Dict[str, Any]]) -> pd.DataFrame:
    """
    Экспорт данных в DataFrame с правильной структурой колонок
    
    Args:
        data: Список записей
        
    Returns:
        pd.DataFrame: Готовый DataFrame
    """
    if not data:
        logger.warning("⚠️ Нет данных для экспорта")
        return pd.DataFrame()
    
    df = pd.DataFrame(data)
    
    # Добавляем отсутствующие колонки как NaN
    for col in EXPECTED_COLUMNS:
        if col not in df.columns:
            df[col] = None
    
    # Строгий порядок колонок + оставшиеся дополнительные (если API вернул новые)
    ordered = [c for c in EXPECTED_COLUMNS if c in df.columns]
    extra = [c for c in df.columns if c not in EXPECTED_COLUMNS]
    df = df[ordered + extra]
    
    return df


def log_statistics(df: pd.DataFrame):
    """
    Логирование статистики заполнения данных
    
    Args:
        df: DataFrame с данными
    """
    if df.empty:
        return
    
    logger.info("📊 Статистика заполнения (ключевые поля):")
    
    key_columns = [
        "guidEs", "Name", "trnNameRus", "regNumber", 
        "mnnRus", "producer", "jnvls"
    ]
    
    for col in key_columns:
        if col in df.columns:
            filled = df[col].notna().sum()
            percentage = filled / len(df) * 100
            logger.info(f"  {col}: {filled}/{len(df)} ({percentage:.1f}%)")


# ============================================================================
# ОСНОВНАЯ ФУНКЦИЯ ЗАГРУЗКИ
# ============================================================================

def download_eprica_data() -> Optional[pd.DataFrame]:
    """
    Загрузка данных справочника ePrica и возврат в виде DataFrame
    
    Returns:
        pd.DataFrame или None если ошибка
    """
    logger.info("=" * 60)
    logger.info("🚀 Запуск выгрузки данных из ePrica...")
    logger.info("=" * 60)
    
    # Шаг 1: Авторизация
    token = get_token()
    if not token:
        return None
    
    # Шаг 2: Полная выгрузка данных с пагинацией
    raw_data = fetch_all_data(token)
    
    if not raw_data:
        logger.error("❌ Не удалось получить данные из API ePrica")
        return None
    
    # Шаг 3: Нормализация и дедупликация
    logger.info("🔧 Нормализация и дедупликация данных...")
    normalized_data = normalize_and_deduplicate(raw_data)
    
    logger.info(f"✅ После дедупликации: {len(normalized_data)} уникальных записей")
    
    # Шаг 4: Создание DataFrame
    df = export_to_dataframe(normalized_data)
    
    if df.empty:
        logger.warning("⚠️ DataFrame пуст после обработки")
        return None
    
    # Шаг 5: Логирование статистики
    log_statistics(df)
    
    logger.info(f"✅ Данные готовы: {len(df)} строк, {len(df.columns)} колонок")
    
    return df


# ============================================================================
# ТОЧКА ВХОДА
# ============================================================================

def main():
    """Основная функция - сохраняет данные в файл (для тестирования)"""
    df = download_eprica_data()
    
    if df is not None:
        output_file = "ePrica_справочник.xlsx"
        
        # Проверка на превышение лимита Excel
        if len(df) > 1_048_576:
            csv_name = output_file.replace(".xlsx", ".csv")
            df.to_csv(csv_name, index=False, encoding="utf-8-sig")
            logger.info(f"💾 Лимит Excel превышен. Сохранено в: {csv_name}")
        else:
            df.to_excel(output_file, index=False, engine="openpyxl")
            logger.info(f"💾 Файл успешно сохранён: {output_file}")
    
    logger.info("🏁 Готово.")


if __name__ == "__main__":
    main()
