#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
pip install -r requirements.txt


Главный скрипт управления процессами сбора справочников.

Функции:
1. Загрузка ИСРАС.xlsx на лист "ИСРАС" файла output/esklp/esklp_full.xlsx
2. Управление запуском остальных скриптов:
   - rls_xml_to_xls.py (парсинг inventory.xml → лист "РЛС")
   - download_grls.py (скачивание ГРЛС → лист "ГРЛС")
   - download_esklp.py (скачивание ЕСКЛП → лист "ЕСКЛП")
3. Загрузка Справочника ЦПЗ API на лист ЦПЗ
4. Загрузка справочника ePrica API на лист "ePrica"

Дата: 2026
"""

import os
import sys
import logging
import subprocess
from pathlib import Path
from datetime import datetime
from typing import Optional, List

import pandas as pd

# ============================================================================
# НАСТРОЙКА ЛОГИРОВАНИЯ
# ============================================================================

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s | %(levelname)-8s | %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler('main.log', encoding='utf-8', mode='w')
    ]
)
logger = logging.getLogger(__name__)


# ============================================================================
# КОНФИГУРАЦИЯ
# ============================================================================

class Config:
    """Конфигурация главного скрипта"""
    
    # Пути к файлам
    NOMEN_FILE = 'nomen.xlsx'
    OUTPUT_DIR = Path('output/esklp')
    ESKLP_FULL_FILE = OUTPUT_DIR / 'esklp_full.xlsx'
    
    # Скрипты для запуска до загрузки ЦПЗ (ЕСКЛП должен быть первым!)
    SCRIPTS_BEFORE_NOMEN = [
        'download_esklp.py',   # Запускается ПЕРВЫМ - создаёт файл с листом ЕСКЛП
    ]
    
    # Скрипты для запуска после загрузки ЦПЗ
    SCRIPTS_AFTER_NOMEN = [
        'rls_xml_to_xls.py',   # Затем РЛС
        'download_grls.py',    # Затем ГРЛС
    ]
    
    # Скрипты для запуска в самом конце (после всех остальных)
    SCRIPTS_FINAL = [
        # ePrica загружается напрямую через import в main.py
    ]
    
    # Листы в esklp_full.xlsx
    SHEET_NOMEN = 'ЦПЗ'
    SHEET_RLS = 'РЛС'
    SHEET_GRLS = 'ГРЛС'
    SHEET_ESKLP = 'ЕСКЛП'
    SHEET_EPRICA = 'ePrica'
    SHEET_ISRAS = 'ИСРАС'

    # Файл ИСРАС
    ISRAS_FILE = 'ИСРАС.xlsx'


# ============================================================================
# КЛАСС УПРАВЛЕНИЯ ПРОЦЕССАМИ
# ============================================================================

class ProcessManager:
    """Управление запуском внешних скриптов и обработкой данных"""
    
    def __init__(self):
        self.output_dir = Config.OUTPUT_DIR
        self.esklp_full_path = Config.ESKLP_FULL_FILE
        
        # Создаём директорию output/esklp если не существует
        self.output_dir.mkdir(parents=True, exist_ok=True)
        
        logger.info(f"📁 Рабочая директория: {self.output_dir.resolve()}")
        logger.info(f"📄 Целевой файл: {self.esklp_full_path}")
    
    def load_nomen_to_sheet(self) -> bool:
        """
        Загрузка файла nomen.xlsx на лист "ЦПЗ" файла esklp_full.xlsx
        
        Returns:
            bool: True если успешно, иначе False
        """
        logger.info("=" * 60)
        logger.info("📋 Этап 1: Загрузка nomen.xlsx на лист 'ЦПЗ'")
        logger.info("=" * 60)
        
        nomen_path = Path(Config.NOMEN_FILE)
        
        # Проверка существования файла nomen.xlsx
        if not nomen_path.exists():
            logger.error(f"❌ Файл {Config.NOMEN_FILE} не найден!")
            return False
        
        try:
            # Чтение nomen.xlsx
            logger.info(f"📖 Чтение файла: {nomen_path}")
            df_nomen = pd.read_excel(nomen_path, engine='openpyxl')
            
            if df_nomen.empty:
                logger.warning(f"⚠️ Файл {Config.NOMEN_FILE} пустой")
            else:
                logger.info(f"✅ Загружено строк: {len(df_nomen)}")
                logger.info(f"✅ Колонки: {list(df_nomen.columns)}")
            
            # Сохранение на лист "ЦПЗ"
            logger.info(f"💾 Сохранение на лист '{Config.SHEET_NOMEN}'...")
            
            if self.esklp_full_path.exists():
                # Файл существует - добавляем/обновляем лист
                with pd.ExcelWriter(
                    self.esklp_full_path,
                    engine='openpyxl',
                    mode='a',
                    if_sheet_exists='replace'
                ) as writer:
                    df_nomen.to_excel(writer, sheet_name=Config.SHEET_NOMEN, index=False)
                logger.info(f"✅ Лист '{Config.SHEET_NOMEN}' обновлён в существующем файле")
            else:
                # Файл не существует - создаём новый
                with pd.ExcelWriter(self.esklp_full_path, engine='openpyxl') as writer:
                    df_nomen.to_excel(writer, sheet_name=Config.SHEET_NOMEN, index=False)
                logger.info(f"✅ Создан новый файл с листом '{Config.SHEET_NOMEN}'")
            
            logger.info(f"✅ Файл сохранён: {self.esklp_full_path}")
            return True
            
        except Exception as e:
            logger.error(f"❌ Ошибка загрузки nomen.xlsx: {type(e).__name__}: {e}")
            return False
    
    def run_script(self, script_name: str) -> bool:
        """
        Запуск указанного скрипта
        
        Args:
            script_name: Имя скрипта для запуска
            
        Returns:
            bool: True если скрипт выполнен успешно (exit code 0)
        """
        script_path = Path(script_name)
        
        if not script_path.exists():
            logger.warning(f"⚠️ Скрипт {script_name} не найден, пропускаем")
            return False
        
        logger.info("-" * 60)
        logger.info(f"▶️  Запуск скрипта: {script_name}")
        logger.info("-" * 60)
        
        try:
            # Запускаем скрипт через subprocess
            result = subprocess.run(
                [sys.executable, str(script_path)],
                cwd=Path.cwd(),
                capture_output=False,
                text=True,
                timeout=1800  # 30 минут таймаут
            )
            
            if result.returncode == 0:
                logger.info(f"✅ Скрипт {script_name} завершён успешно")
                return True
            else:
                logger.error(f"❌ Скрипт {script_name} завершился с кодом {result.returncode}")
                return False
                
        except subprocess.TimeoutExpired:
            logger.error(f"❌ Таймаут выполнения скрипта {script_name} (30 мин)")
            return False
        except Exception as e:
            logger.error(f"❌ Ошибка запуска {script_name}: {type(e).__name__}: {e}")
            return False
    
    def run_all_scripts(self, scripts_list: list) -> dict:
        """
        Последовательный запуск всех скриптов из списка
        
        Args:
            scripts_list: Список скриптов для запуска
            
        Returns:
            dict: Результаты выполнения каждого скрипта
        """
        results = {}
        
        for script in scripts_list:
            success = self.run_script(script)
            results[script] = success
            
            # Небольшая пауза между скриптами
            if script != scripts_list[-1]:
                logger.info("⏳ Пауза 2 секунды перед следующим скриптом...")
                import time
                time.sleep(2)
        
        return results
    
    def verify_result(self) -> dict:
        """
        Проверка результата - какие листы существуют в файле
        
        Returns:
            dict: Информация о листах в файле
        """
        logger.info("=" * 60)
        logger.info("🔍 Этап 3: Проверка результата")
        logger.info("=" * 60)
        
        result_info = {
            'file_exists': False,
            'sheets': [],
            'sheet_stats': {}
        }
        
        if not self.esklp_full_path.exists():
            logger.error(f"❌ Файл {self.esklp_full_path} не создан")
            return result_info
        
        result_info['file_exists'] = True
        
        try:
            # Получаем список всех листов
            xl = pd.ExcelFile(self.esklp_full_path, engine='openpyxl')
            result_info['sheets'] = xl.sheet_names
            
            logger.info(f"📄 Файл: {self.esklp_full_path}")
            logger.info(f"📑 Листы в файле: {xl.sheet_names}")
            
            # Статистика по каждому листу
            for sheet in xl.sheet_names:
                df = pd.read_excel(xl, sheet_name=sheet, nrows=0)
                row_count = pd.read_excel(xl, sheet_name=sheet).shape[0]
                col_count = len(df.columns)
                result_info['sheet_stats'][sheet] = {
                    'rows': row_count,
                    'columns': col_count
                }
                logger.info(f"   • {sheet}: {row_count} строк × {col_count} колонок")
            
            xl.close()
            
        except Exception as e:
            logger.error(f"❌ Ошибка проверки файла: {type(e).__name__}: {e}")
        
        return result_info


    def load_eprica_to_sheet(self) -> bool:
        """
        Загрузка справочника ePrica напрямую на лист "ePrica" файла esklp_full.xlsx
        (без промежуточного Excel-файла)

        Returns:
            bool: True если успешно, иначе False
        """
        logger.info("=" * 60)
        logger.info("📋 Этап: Загрузка справочника ePrica на лист 'ePrica'")
        logger.info("=" * 60)

        try:
            # Импортируем функцию загрузки из download_eprica.py
            from download_eprica import download_eprica_data
            
            # Получаем данные напрямую в DataFrame
            logger.info("📥 Загрузка данных из API ePrica...")
            df_eprica = download_eprica_data()

            if df_eprica is None or df_eprica.empty:
                logger.error("❌ Не удалось получить данные из API ePrica")
                return False

            logger.info(f"✅ Загружено строк: {len(df_eprica)}")
            logger.info(f"✅ Колонки: {list(df_eprica.columns)}")

            # Сохранение на лист "ePrica"
            logger.info(f"💾 Сохранение на лист '{Config.SHEET_EPRICA}'...")

            if self.esklp_full_path.exists():
                # Файл существует - добавляем/обновляем лист
                with pd.ExcelWriter(
                    self.esklp_full_path,
                    engine='openpyxl',
                    mode='a',
                    if_sheet_exists='replace'
                ) as writer:
                    df_eprica.to_excel(writer, sheet_name=Config.SHEET_EPRICA, index=False)
                logger.info(f"✅ Лист '{Config.SHEET_EPRICA}' обновлён в существующем файле")
            else:
                # Файл не существует - создаём новый
                with pd.ExcelWriter(self.esklp_full_path, engine='openpyxl') as writer:
                    df_eprica.to_excel(writer, sheet_name=Config.SHEET_EPRICA, index=False)
                logger.info(f"✅ Создан новый файл с листом '{Config.SHEET_EPRICA}'")

            logger.info(f"✅ Файл сохранён: {self.esklp_full_path}")
            return True

        except ImportError as e:
            logger.error(f"❌ Ошибка импорта download_eprica: {e}")
            return False
        except Exception as e:
            logger.error(f"❌ Ошибка загрузки справочника ePrica: {type(e).__name__}: {e}")
            return False

    def load_isras_to_sheet(self) -> bool:
        """
        Загрузка файла ИСРАС.xlsx на лист "ИСРАС" файла esklp_full.xlsx

        Returns:
            bool: True если успешно, иначе False
        """
        logger.info("=" * 60)
        logger.info("📋 Этап: Загрузка файла ИСРАС.xlsx на лист 'ИСРАС'")
        logger.info("=" * 60)

        isras_path = Path(Config.ISRAS_FILE)

        # Проверка существования файла ИСРАС.xlsx
        if not isras_path.exists():
            logger.error(f"❌ Файл {Config.ISRAS_FILE} не найден!")
            return False

        try:
            # Чтение файла ИСРАС.xlsx
            logger.info(f"📖 Чтение файла: {isras_path}")
            df_isras = pd.read_excel(isras_path, engine='openpyxl')

            if df_isras.empty:
                logger.warning(f"⚠️ Файл {Config.ISRAS_FILE} пустой")
            else:
                logger.info(f"✅ Загружено строк: {len(df_isras)}")
                logger.info(f"✅ Колонки: {list(df_isras.columns)}")

            # Сохранение на лист "ИСРАС"
            logger.info(f"💾 Сохранение на лист '{Config.SHEET_ISRAS}'...")

            if self.esklp_full_path.exists():
                # Файл существует - добавляем/обновляем лист
                with pd.ExcelWriter(
                    self.esklp_full_path,
                    engine='openpyxl',
                    mode='a',
                    if_sheet_exists='replace'
                ) as writer:
                    df_isras.to_excel(writer, sheet_name=Config.SHEET_ISRAS, index=False)
                logger.info(f"✅ Лист '{Config.SHEET_ISRAS}' обновлён в существующем файле")
            else:
                # Файл не существует - создаём новый
                with pd.ExcelWriter(self.esklp_full_path, engine='openpyxl') as writer:
                    df_isras.to_excel(writer, sheet_name=Config.SHEET_ISRAS, index=False)
                logger.info(f"✅ Создан новый файл с листом '{Config.SHEET_ISRAS}'")

            logger.info(f"✅ Файл сохранён: {self.esklp_full_path}")
            return True

        except Exception as e:
            logger.error(f"❌ Ошибка загрузки ИСРАС.xlsx: {type(e).__name__}: {e}")
            return False

    def run_annotation_analysis(self) -> bool:
        """
        Запуск скрипта create_annotation.py для создания листа "Аннотация"
        с анализом связей между таблицами
        
        Returns:
            bool: True если скрипт выполнен успешно
        """
        logger.info("=" * 60)
        logger.info("📋 Этап: Запуск анализа связей (create_annotation.py)")
        logger.info("=" * 60)
        
        script_path = Path('create_annotation.py')
        
        if not script_path.exists():
            logger.warning(f"⚠️ Скрипт {script_path} не найден, пропускаем")
            return False
        
        try:
            # Запускаем скрипт через subprocess
            result = subprocess.run(
                [sys.executable, str(script_path)],
                cwd=Path.cwd(),
                capture_output=False,
                text=True,
                timeout=1800  # 30 минут таймаут
            )
            
            if result.returncode == 0:
                logger.info(f"✅ Скрипт {script_path} завершён успешно")
                logger.info("✅ Лист 'Аннотация' добавлен в файл esklp_full.xlsx")
                return True
            else:
                logger.error(f"❌ Скрипт {script_path} завершился с кодом {result.returncode}")
                return False
                
        except subprocess.TimeoutExpired:
            logger.error(f"❌ Таймаут выполнения скрипта {script_path} (30 мин)")
            return False
        except Exception as e:
            logger.error(f"❌ Ошибка запуска {script_path}: {type(e).__name__}: {e}")
            return False


# ============================================================================
# ОСНОВНАЯ ФУНКЦИЯ
# ============================================================================

def main():
    """Точка входа программы"""
    
    logger.info("=" * 60)
    logger.info("🎯 ЗАПУСК ГЛАВНОГО СКРИПТА ЕСКЛП")
    logger.info(f"📅 Дата: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    logger.info("=" * 60)
    
    manager = ProcessManager()
    
    # Этап 1: Сначала запускаем download_esklp.py (создаёт файл с листом ЕСКЛП)
    logger.info("=" * 60)
    logger.info("🚀 Этап 1: Запуск download_esklp.py (создание листа ЕСКЛП)")
    logger.info("=" * 60)
    script_results_before = manager.run_all_scripts(Config.SCRIPTS_BEFORE_NOMEN)
    
    # Этап 2: Загрузка nomen.xlsx на лист "ЦПЗ"
    nomen_success = manager.load_nomen_to_sheet()
    
    if not nomen_success:
        logger.error("❌ Критическая ошибка: не удалось загрузить nomen.xlsx")
        logger.info("⏭️  Продолжаем выполнение без загрузки ЦПЗ...")
    
    # Этап 3: Запуск остальных скриптов (РЛС и ГРЛС)
    logger.info("=" * 60)
    logger.info("🚀 Этап 3: Запуск остальных скриптов (РЛС, ГРЛС)")
    logger.info("=" * 60)
    script_results_after = manager.run_all_scripts(Config.SCRIPTS_AFTER_NOMEN)
    
    # Этап 4: Загрузка справочника ePrica (последним!)
    logger.info("=" * 60)
    logger.info("🚀 Этап 4: Загрузка справочника ePrica (лист 'ePrica')")
    logger.info("=" * 60)
    eprica_success = manager.load_eprica_to_sheet()
    
    if not eprica_success:
        logger.error("❌ Ошибка: не удалось загрузить справочник ePrica")
    
    # Этап 5: Загрузка файла ИСРАС.xlsx на лист "ИСРАС"
    logger.info("=" * 60)
    logger.info("🚀 Этап 5: Загрузка файла ИСРАС.xlsx (лист 'ИСРАС')")
    logger.info("=" * 60)
    isras_success = manager.load_isras_to_sheet()
    
    if not isras_success:
        logger.error("❌ Ошибка: не удалось загрузить файл ИСРАС.xlsx")
    
    # Этап 6: Запуск анализа связей (create_annotation.py)
    logger.info("=" * 60)
    logger.info("🚀 Этап 6: Запуск анализа связей (create_annotation.py)")
    logger.info("=" * 60)
    annotation_success = manager.run_annotation_analysis()
    
    if not annotation_success:
        logger.error("❌ Ошибка: не удалось выполнить анализ связей")
    
    # Объединяем результаты всех скриптов
    script_results = {**script_results_before, **script_results_after}
    
    # Этап 7: Проверка результата
    result_info = manager.verify_result()
    
    # Итоговый отчёт
    logger.info("=" * 60)
    logger.info("📊 ИТОГОВЫЙ ОТЧЁТ")
    logger.info("=" * 60)
    
    # Статус загрузки nomen
    logger.info(f"{'✅' if nomen_success else '❌'} Загрузка nomen.xlsx (лист ЦПЗ)")
    
    # Статус загрузки ePrica
    logger.info(f"{'✅' if eprica_success else '❌'} Загрузка справочника ePrica (лист ePrica)")
    
    # Статус загрузки ИСРАС
    logger.info(f"{'✅' if isras_success else '❌'} Загрузка ИСРАС.xlsx (лист ИСРАС)")
    
    # Статус анализа аннотаций
    logger.info(f"{'✅' if annotation_success else '❌'} Анализ связей (лист Аннотация)")
    
    # Статус скриптов
    for script, success in script_results.items():
        status = '✅' if success else '❌'
        logger.info(f"{status} Скрипт {script}")
    
    # Информация о файле
    if result_info['file_exists']:
        logger.info(f"✅ Файл создан: {Config.ESKLP_FULL_FILE}")
        logger.info(f"📑 Листы: {result_info['sheets']}")
    else:
        logger.warning(f"⚠️ Файл {Config.ESKLP_FULL_FILE} не найден")
    
    logger.info("=" * 60)
    logger.info("✅ РАБОТА ЗАВЕРШЕНА")
    logger.info("=" * 60)
    
    return {
        'nomen_loaded': nomen_success,
        'eprica_loaded': eprica_success,
        'isras_loaded': isras_success,
        'annotation_analyzed': annotation_success,
        'scripts': script_results,
        'file_info': result_info
    }


# ============================================================================
# ЗАПУСК
# ============================================================================

if __name__ == '__main__':
    try:
        result = main()
        
        # Возвращаем код выхода в зависимости от успеха
        if result['nomen_loaded'] and result.get('eprica_loaded', True) and all(result['scripts'].values()):
            sys.exit(0)  # Успех
        else:
            sys.exit(1)  # Частичный успех или ошибка
            
    except KeyboardInterrupt:
        logger.warning("\n⚠️ Прервано пользователем")
        sys.exit(130)
    except Exception as e:
        logger.error(f"❌ Критическая ошибка: {type(e).__name__}: {e}", exc_info=True)
        sys.exit(2)
