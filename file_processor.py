"""
Модуль обработки и парсинга файлов для Джарвиса.
Поддерживает: Excel (.xlsx, .xls), CSV, PDF, Word (.docx), текстовые форматы.
"""

import io
import json
import logging
from pathlib import Path
from typing import Any, Dict, List, Optional

import pandas as pd
import openpyxl
from pypdf import PdfReader
from docx import Document

logger = logging.getLogger("JarvisFileProcessor")


def format_dataframe_summary(df: pd.DataFrame, max_rows: int = 15) -> str:
    """Генерация детальной сводки по таблице DataFrame."""
    lines = []
    rows, cols = df.shape
    lines.append(f"- Размер таблицы: {rows} строк, {cols} колонок")
    lines.append(f"- Колонки: {', '.join([str(c) for c in df.columns[:30]])}")
    
    # Расчет сумм по числовым колонкам (для финансов, выручки, расходов)
    numeric_cols = df.select_dtypes(include=["number"]).columns
    if len(numeric_cols) > 0:
        lines.append("\n[Итоговые суммы по числовым колонкам]")
        for col in numeric_cols[:20]:
            total = df[col].sum()
            lines.append(f"  * {col}: {total:,.2f}".replace(",", " "))

    # Пример данных
    lines.append("\n[Образец данных - первые строки]")
    preview = df.head(max_rows)
    lines.append(preview.to_string(index=False, max_cols=12))

    return "\n".join(lines)


def parse_excel(file_path: Path) -> str:
    """Парсинг Excel файлов (.xlsx, .xls) со всеми листами."""
    output = [f"[Файл Excel: {file_path.name}]"]
    try:
        with pd.ExcelFile(file_path) as excel_file:
            sheet_names = excel_file.sheet_names
            output.append(f"Листов в файле: {len(sheet_names)} ({', '.join(sheet_names)})")

            for sheet in sheet_names:
                output.append(f"\n--- Лист: {sheet} ---")
                try:
                    df = pd.read_excel(excel_file, sheet_name=sheet)
                    if df.empty:
                        output.append("Лист пустой.")
                        continue
                    
                    # Удаляем полностью пустые строки и колонки
                    df = df.dropna(how="all").dropna(axis=1, how="all")
                    
                    # Если строк немного (< 35), выводим всю таблицу
                    if len(df) <= 35:
                        output.append(f"Размер: {df.shape[0]} строк x {df.shape[1]} колонок")
                        output.append(df.to_string(index=False, max_cols=15))
                    else:
                        output.append(format_dataframe_summary(df, max_rows=10))
                except Exception as e:
                    output.append(f"Ошибка чтения листа {sheet}: {e}")
    except Exception as e:
        output.append(f"Ошибка открытия файла Excel: {e}")

    return "\n".join(output)


def parse_csv(file_path: Path) -> str:
    """Парсинг CSV файлов с автоопределением кодировки и разделителя."""
    output = [f"[CSV Файл: {file_path.name}]"]
    
    # Пробуем разные кодировки (UTF-8, CP1251)
    df = None
    for enc in ["utf-8", "cp1251", "latin1"]:
        for sep in [",", ";", "\t"]:
            try:
                temp_df = pd.read_csv(file_path, encoding=enc, sep=sep, nrows=5)
                if temp_df.shape[1] > 1:
                    df = pd.read_csv(file_path, encoding=enc, sep=sep)
                    break
            except Exception:
                continue
        if df is not None:
            break

    if df is None:
        try:
            df = pd.read_csv(file_path)
        except Exception as e:
            return f"Ошибка чтения CSV: {e}"

    if len(df) <= 35:
        output.append(df.to_string(index=False))
    else:
        output.append(format_dataframe_summary(df, max_rows=12))
    return "\n".join(output)


def parse_pdf(file_path: Path, max_pages: int = 25) -> str:
    """Парсинг PDF файлов."""
    output = [f"[PDF Документ: {file_path.name}]"]
    try:
        reader = PdfReader(str(file_path))
        num_pages = len(reader.pages)
        output.append(f"Всего страниц: {num_pages}")
        
        pages_to_read = min(num_pages, max_pages)
        text_content = []
        for i in range(pages_to_read):
            page_text = reader.pages[i].extract_text() or ""
            if page_text.strip():
                text_content.append(f"[Страница {i+1}]\n{page_text.strip()}")
                
        output.append("\n\n".join(text_content))
        if num_pages > max_pages:
            output.append(f"\n...(пропущено страниц: {num_pages - max_pages})")
    except Exception as e:
        output.append(f"Ошибка чтения PDF: {e}")
        
    return "\n".join(output)


def parse_docx(file_path: Path) -> str:
    """Парсинг Word документов (.docx)."""
    output = [f"[Word Документ: {file_path.name}]"]
    try:
        doc = Document(str(file_path))
        paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
        output.append("Текст документа:")
        output.append("\n".join(paragraphs[:100]))
        
        # Парсинг таблиц в документе
        if doc.tables:
            output.append(f"\nНайдено таблиц: {len(doc.tables)}")
            for idx, table in enumerate(doc.tables[:5]):
                output.append(f"\nТаблица {idx+1}:")
                for row in table.rows[:15]:
                    row_data = [cell.text.strip() for cell in row.cells]
                    output.append(" | ".join(row_data))
    except Exception as e:
        output.append(f"Ошибка чтения Word: {e}")
        
    return "\n".join(output)


def parse_text_file(file_path: Path, max_chars: int = 25000) -> str:
    """Чтение текстовых файлов (txt, py, json, md, log)."""
    for enc in ["utf-8", "utf-8-sig", "cp1251", "latin1"]:
        try:
            content = file_path.read_text(encoding=enc)
            if len(content) > max_chars:
                return f"[Текстовый файл: {file_path.name}]\n\n{content[:max_chars]}\n\n...(файл обрезан до {max_chars} символов)"
            return f"[Текстовый файл: {file_path.name}]\n\n{content}"
        except UnicodeDecodeError:
            continue
    return f"Не удалось прочитать текстовый файл {file_path.name} в поддерживаемых кодировках."


def extract_file_content(file_path: Path) -> str:
    """Главная функция: определяет тип файла и возвращает структурированный текст."""
    suffix = file_path.suffix.lower()
    
    if suffix in [".xlsx", ".xls", ".xlsm"]:
        return parse_excel(file_path)
    elif suffix in [".csv", ".tsv"]:
        return parse_csv(file_path)
    elif suffix == ".pdf":
        return parse_pdf(file_path)
    elif suffix in [".docx", ".doc"]:
        return parse_docx(file_path)
    elif suffix in [".txt", ".py", ".json", ".xml", ".html", ".md", ".log", ".yaml", ".yml", ".sql"]:
        return parse_text_file(file_path)
    else:
        # Пробуем как текст
        try:
            return parse_text_file(file_path)
        except Exception:
            return f"Файл {file_path.name} имеет неподдерживаемый формат ({suffix})."
