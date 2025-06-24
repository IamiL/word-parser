"""
Модуль для обработки DOCX файлов и извлечения текста
"""

import re
from docx import Document
from io import BytesIO


def process_docx_file(file_content):
    """
    Обрабатывает DOCX файл и извлекает текст, убирая спецсимволы
    
    Args:
        file_content (bytes): Содержимое DOCX файла в байтах
        
    Returns:
        str: Очищенный текст с сохранением абзацев
        
    Raises:
        Exception: При ошибке обработки файла
    """
    try:
        # Создаем BytesIO объект из содержимого файла
        file_stream = BytesIO(file_content)
        
        # Открываем документ
        document = Document(file_stream)
        
        # Извлекаем текст из всех параграфов
        paragraphs = []
        for paragraph in document.paragraphs:
            text = paragraph.text.strip()
            if text:  # Добавляем только непустые параграфы
                # Очищаем текст от спецсимволов, оставляя только буквы, цифры, пунктуацию и пробелы
                cleaned_text = clean_text(text)
                if cleaned_text:
                    paragraphs.append(cleaned_text)
        
        # Соединяем параграфы через двойной перенос строки
        result = '\n\n'.join(paragraphs)
        
        return result
        
    except Exception as e:
        raise Exception(f"Ошибка при обработке DOCX файла: {str(e)}")


def clean_text(text):
    """
    Очищает текст от спецсимволов, оставляя только читаемые символы
    
    Args:
        text (str): Исходный текст
        
    Returns:
        str: Очищенный текст
    """
    # Убираем управляющие символы, кроме переносов строк и табуляций
    # Оставляем буквы, цифры, пунктуацию, пробелы
    cleaned = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F]', '', text)
    
    # Нормализуем пробелы - заменяем множественные пробелы на одинарные
    cleaned = re.sub(r'\s+', ' ', cleaned)
    
    # Убираем пробелы в начале и конце
    cleaned = cleaned.strip()
    
    return cleaned


def extract_text_from_tables(document):
    """
    Извлекает текст из таблиц документа (дополнительная функция)
    
    Args:
        document: Объект документа python-docx
        
    Returns:
        str: Текст из таблиц
    """
    table_texts = []
    
    for table in document.tables:
        table_content = []
        for row in table.rows:
            row_content = []
            for cell in row.cells:
                cell_text = cell.text.strip()
                if cell_text:
                    cleaned_cell = clean_text(cell_text)
                    if cleaned_cell:
                        row_content.append(cleaned_cell)
            
            if row_content:
                table_content.append(' | '.join(row_content))
        
        if table_content:
            table_texts.append('\n'.join(table_content))
    
    return '\n\n'.join(table_texts)


def process_docx_with_tables(file_content):
    """
    Расширенная версия обработки DOCX с извлечением текста из таблиц
    
    Args:
        file_content (bytes): Содержимое DOCX файла в байтах
        
    Returns:
        str: Полный текст документа включая таблицы
    """
    try:
        file_stream = BytesIO(file_content)
        document = Document(file_stream)
        
        # Извлекаем основной текст
        main_text = []
        for paragraph in document.paragraphs:
            text = paragraph.text.strip()
            if text:
                cleaned_text = clean_text(text)
                if cleaned_text:
                    main_text.append(cleaned_text)
        
        # Извлекаем текст из таблиц
        table_text = extract_text_from_tables(document)
        
        # Объединяем результаты
        result_parts = []
        if main_text:
            result_parts.append('\n\n'.join(main_text))
        
        if table_text:
            result_parts.append(table_text)
        
        return '\n\n---\n\n'.join(result_parts)
        
    except Exception as e:
        raise Exception(f"Ошибка при обработке DOCX файла: {str(e)}")
