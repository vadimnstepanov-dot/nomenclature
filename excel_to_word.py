#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт для преобразования отчетов из Excel в Word
Анализирует файл esklp_full.xlsx и создает документы Word для каждого листа
"""

import os
import sys
import logging
from pathlib import Path
from datetime import datetime

import pandas as pd
from docx import Document
from docx.shared import Inches, Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('excel_to_word.log', encoding='utf-8', mode='w'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


def set_cell_background(cell, color):
    """Установка цвета фона ячейки"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = tcPr.find(qn('w:shd'))
    if shd is None:
        from docx.oxml import OxmlElement
        shd = OxmlElement('w:shd')
        tcPr.append(shd)
    shd.set(qn('w:fill'), color)


def create_word_report_from_excel(input_file=None, output_dir=None, mode='single'):
    """
    Создание отчетов Word из файла Excel
    
    Args:
        input_file: Путь к файлу Excel
        output_dir: Директория для сохранения отчетов Word
        mode: 'single' - все листы в одном документе, 'separate' - каждый лист в отдельном файле
    
    Returns:
        dict: Информация о созданных файлах
    """
    # Пути по умолчанию
    if input_file is None:
        input_file = os.path.join(os.getcwd(), 'output', 'esklp', 'esklp_full.xlsx')
    
    if output_dir is None:
        output_dir = os.path.join(os.getcwd(), 'output', 'esklp', 'word_reports')
    
    # Создаем директорию для отчетов
    Path(output_dir).mkdir(parents=True, exist_ok=True)
    
    logger.info(f"Загрузка файла {input_file}")
    
    # Проверяем существование файла
    if not os.path.exists(input_file):
        logger.error(f"Файл {input_file} не найден!")
        return {'success': False, 'error': 'File not found'}
    
    try:
        # Загружаем все листы
        xls = pd.ExcelFile(input_file)
        sheet_names = xls.sheet_names
        logger.info(f"Найдены листы: {sheet_names}")
        
        created_files = []
        
        if mode == 'separate':
            # Каждый лист в отдельном документе
            for sheet_name in sheet_names:
                logger.info(f"Обработка листа: {sheet_name}")
                
                # Загружаем данные листа
                df = pd.read_excel(input_file, sheet_name=sheet_name)
                logger.info(f"  Загружено строк: {len(df)}, колонок: {len(df.columns)}")
                
                # Создаем документ Word
                doc = Document()
                
                # Добавляем заголовок
                heading = doc.add_heading(sheet_name, level=1)
                heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # Добавляем информацию о дате формирования
                timestamp_para = doc.add_paragraph()
                timestamp_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = timestamp_para.add_run(f"Дата формирования: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
                run.italic = True
                
                doc.add_paragraph()  # Пустая строка
                
                # Добавляем статистику
                stats_para = doc.add_heading("Статистика", level=2)
                stats_info = [
                    f"Всего записей: {len(df)}",
                    f"Всего колонок: {len(df.columns)}",
                    f"Полностью заполненных колонок: {sum(1 for col in df.columns if df[col].notna().all())}",
                    f"Колонок с пропусками: {sum(1 for col in df.columns if df[col].isna().any())}"
                ]
                for stat in stats_info:
                    doc.add_paragraph(stat, style='List Bullet')
                
                doc.add_paragraph()  # Пустая строка
                
                # Таблица с данными (ограничиваем количество строк для Word)
                max_rows = 1000  # Ограничение для производительности Word
                display_df = df.head(max_rows)
                
                heading_data = doc.add_heading(f"Данные (выборка: {len(display_df)} записей)", level=2)
                
                # Для больших таблиц создаем упрощенный вариант - только первые колонки
                max_cols_to_display = min(15, len(display_df.columns))  # Показываем не более 15 колонок
                
                if len(display_df.columns) > max_cols_to_display:
                    info_para = doc.add_paragraph()
                    info_run = info_para.add_run(
                        f"ℹ️ Показаны первые {max_cols_to_display} колонок из {len(display_df.columns)}. "
                    )
                    info_run.italic = True
                
                # Создаем таблицу
                table = doc.add_table(rows=1, cols=max_cols_to_display)
                table.style = 'Table Grid'
                table.alignment = WD_TABLE_ALIGNMENT.CENTER
                
                # Заполняем заголовки
                header_row = table.rows[0]
                for i in range(max_cols_to_display):
                    col_name = display_df.columns[i]
                    cell = header_row.cells[i]
                    cell.text = str(col_name)[:25]  # Ограничиваем длину
                    # Делаем текст жирным
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            run.bold = True
                            run.font.size = Pt(8)
                    # Устанавливаем фон заголовка
                    set_cell_background(cell, '4472C4')  # Синий цвет
                
                # Заполняем данными
                for _, row in display_df.iterrows():
                    row_cells = table.add_row().cells
                    for i in range(max_cols_to_display):
                        value = row.iloc[i]
                        cell = row_cells[i]
                        cell.text = str(value)[:50] if pd.notna(value) else ''  # Ограничиваем длину
                        for paragraph in cell.paragraphs:
                            for run in paragraph.runs:
                                run.font.size = Pt(7)
                
                doc.add_paragraph()  # Пустая строка
                
                # Добавляем информацию об ограничении
                if len(df) > max_rows or len(df.columns) > max_cols_to_display:
                    warning_para = doc.add_paragraph()
                    parts = []
                    if len(df) > max_rows:
                        parts.append(f"Показаны только первые {max_rows} записей из {len(df)}")
                    if len(df.columns) > max_cols_to_display:
                        parts.append(f"показаны первые {max_cols_to_display} колонок из {len(df.columns)}")
                    warning_run = warning_para.add_run(
                        f"⚠️ {'; '.join(parts)}. Полные данные доступны в исходном Excel файле."
                    )
                    warning_run.italic = True
                
                # Сохраняем документ
                safe_sheet_name = "".join(c for c in sheet_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
                output_file = os.path.join(output_dir, f"{safe_sheet_name}.docx")
                doc.save(output_file)
                logger.info(f"  Сохранено в: {output_file}")
                created_files.append(output_file)
            
            logger.info(f"Создано {len(created_files)} файлов Word")
            
        elif mode == 'single':
            # Все листы в одном документе
            doc = Document()
            
            # Главный заголовок
            main_heading = doc.add_heading("Отчет по справочникам лекарственных препаратов", level=1)
            main_heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # Дата формирования
            timestamp_para = doc.add_paragraph()
            timestamp_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = timestamp_para.add_run(f"Дата формирования: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            run.italic = True
            
            doc.add_page_break()
            
            # Оглавление
            toc_heading = doc.add_heading("Оглавление", level=2)
            for i, sheet_name in enumerate(sheet_names, 1):
                toc_para = doc.add_paragraph()
                toc_run = toc_para.add_run(f"{i}. {sheet_name}")
                toc_run.bold = True
            
            doc.add_page_break()
            
            # Обработка каждого листа
            for idx, sheet_name in enumerate(sheet_names, 1):
                logger.info(f"Обработка листа: {sheet_name}")
                
                # Заголовок раздела
                section_heading = doc.add_heading(f"{idx}. {sheet_name}", level=1)
                section_heading.page_break_before = True
                
                # Загружаем данные листа
                df = pd.read_excel(input_file, sheet_name=sheet_name)
                logger.info(f"  Загружено строк: {len(df)}, колонок: {len(df.columns)}")
                
                # Статистика
                stats_heading = doc.add_heading("Статистика", level=2)
                stats_info = [
                    f"Всего записей: {len(df):,}",
                    f"Всего колонок: {len(df.columns)}",
                    f"Полностью заполненных колонок: {sum(1 for col in df.columns if df[col].notna().all())}",
                    f"Колонок с пропусками: {sum(1 for col in df.columns if df[col].isna().any())}"
                ]
                for stat in stats_info:
                    doc.add_paragraph(stat, style='List Bullet')
                
                # Детальная статистика по колонкам (только первые 20)
                stats_detail_heading = doc.add_heading("Заполненность по колонкам (первые 20)", level=3)
                stats_table = doc.add_table(rows=1, cols=4)
                stats_table.style = 'Table Grid'
                
                # Заголовки таблицы статистики
                stats_headers = ['Колонка', 'Заполнено', 'Пусто', '% заполнения']
                header_row = stats_table.rows[0]
                for i, header in enumerate(stats_headers):
                    cell = header_row.cells[i]
                    cell.text = header
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            run.bold = True
                
                # Данные статистики (ограничиваем первыми 20 колонками)
                max_stats_cols = min(20, len(df.columns))
                for col in list(df.columns)[:max_stats_cols]:
                    non_null_count = df[col].notna().sum()
                    null_count = len(df) - non_null_count
                    fill_percent = round(non_null_count / len(df) * 100, 2) if len(df) > 0 else 0
                    
                    row_cells = stats_table.add_row().cells
                    row_cells[0].text = str(col)[:50]  # Ограничиваем длину названия
                    row_cells[1].text = str(non_null_count)
                    row_cells[2].text = str(null_count)
                    row_cells[3].text = f"{fill_percent}%"
                
                if len(df.columns) > max_stats_cols:
                    info_para = doc.add_paragraph()
                    info_run = info_para.add_run(f"ℹ️ Показаны первые {max_stats_cols} колонок из {len(df.columns)}")
                    info_run.italic = True
                
                doc.add_paragraph()  # Пустая строка
                
                # Таблица с данными (ограничиваем для single режима еще сильнее)
                max_rows = 500  # Ограничение для производительности в single режиме
                display_df = df.head(max_rows)
                
                # Для больших таблиц создаем упрощенный вариант - только первые колонки
                max_cols_to_display = min(10, len(display_df.columns))  # Показываем не более 10 колонок
                
                data_heading = doc.add_heading(f"Данные (выборка: {len(display_df):,} записей, {max_cols_to_display} колонок)", level=2)
                
                if len(df.columns) > max_cols_to_display:
                    info_para = doc.add_paragraph()
                    info_run = info_para.add_run(
                        f"ℹ️ Показаны первые {max_cols_to_display} колонок из {len(df.columns)}. "
                    )
                    info_run.italic = True
                
                # Создаем таблицу с данными
                table = doc.add_table(rows=1, cols=max_cols_to_display)
                table.style = 'Table Grid'
                
                # Настраиваем ширину колонок
                for col_idx in range(max_cols_to_display):
                    table.columns[col_idx].width = Cm(3)
                
                # Заполняем заголовки
                header_row = table.rows[0]
                for i in range(max_cols_to_display):
                    col_name = display_df.columns[i]
                    cell = header_row.cells[i]
                    cell.text = str(col_name)[:20]  # Ограничиваем длину
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            run.bold = True
                            run.font.size = Pt(7)
                    set_cell_background(cell, '4472C4')
                
                # Заполняем данными
                for _, row in display_df.iterrows():
                    row_cells = table.add_row().cells
                    for i in range(max_cols_to_display):
                        value = row.iloc[i]
                        cell = row_cells[i]
                        cell.text = str(value)[:50] if pd.notna(value) else ''  # Ограничиваем длину текста
                        for paragraph in cell.paragraphs:
                            for run in paragraph.runs:
                                run.font.size = Pt(6)
                
                # Предупреждение об ограничении
                if len(df) > max_rows:
                    warning_para = doc.add_paragraph()
                    warning_run = warning_para.add_run(
                        f"⚠️ Показаны только первые {max_rows:,} записей из {len(df):,}. "
                        f"Полные данные доступны в исходном Excel файле."
                    )
                    warning_run.italic = True
                
                doc.add_page_break()
            
            # Сохраняем единый документ
            output_file = os.path.join(output_dir, "esklp_full_report.docx")
            doc.save(output_file)
            logger.info(f"Сохранено в: {output_file}")
            created_files.append(output_file)
        
        return {
            'success': True,
            'files': created_files,
            'mode': mode,
            'sheets_processed': len(sheet_names)
        }
        
    except Exception as e:
        logger.error(f"Ошибка при создании отчетов: {type(e).__name__}: {e}")
        import traceback
        logger.error(traceback.format_exc())
        return {'success': False, 'error': str(e)}


if __name__ == '__main__':
    import argparse
    
    parser = argparse.ArgumentParser(description='Преобразование Excel отчетов в Word')
    parser.add_argument('--input', '-i', help='Путь к входному файлу Excel')
    parser.add_argument('--output', '-o', help='Директория для выходных файлов Word')
    parser.add_argument('--mode', '-m', choices=['single', 'separate'], default='single',
                       help='Режим: single - все в одном файле, separate - каждый лист отдельно')
    
    args = parser.parse_args()
    
    result = create_word_report_from_excel(
        input_file=args.input,
        output_dir=args.output,
        mode=args.mode
    )
    
    if result['success']:
        print(f"\n✓ Успешно создано {len(result['files'])} файлов Word")
        for f in result['files']:
            print(f"  - {f}")
    else:
        print(f"\n✗ Ошибка: {result.get('error', 'Неизвестная ошибка')}")
        sys.exit(1)
