"""
HTTP сервер для обработки DOCX файлов
"""

from flask import Flask, request, jsonify, make_response
from werkzeug.exceptions import BadRequest, RequestEntityTooLarge
import logging
from docx_processor import process_docx_file


# Конфигурация приложения
app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # Максимальный размер файла 50MB

# Настройка логирования
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


@app.route('/', methods=['GET'])
def index():
    """
    Главная страница с информацией об API
    """
    return jsonify({
        "service": "DOCX to Text Converter",
        "version": "1.0",
        "endpoints": {
            "convert": {
                "url": "/convert",
                "method": "POST",
                "description": "Конвертирует DOCX файл в текст",
                "content_type": "multipart/form-data",
                "file_field": "file"
            },
            "convert_with_tables": {
                "url": "/convert-with-tables", 
                "method": "POST",
                "description": "Конвертирует DOCX файл в текст включая таблицы",
                "content_type": "multipart/form-data",
                "file_field": "file"
            }
        },
        "usage": "curl -X POST -F 'file=@document.docx' http://localhost:port/convert"
    })


@app.route('/convert', methods=['POST'])
def convert_docx():
    """
    Обработчик для конвертации DOCX файла в текст
    
    Returns:
        JSON с результатом или ошибкой
    """
    try:
        # Проверяем наличие файла в запросе
        if 'file' not in request.files:
            return jsonify({'error': 'Файл не найден в запросе. Используйте поле "file"'}), 400
            
        file = request.files['file']
        
        # Проверяем, что файл был выбран
        if file.filename == '':
            return jsonify({'error': 'Файл не выбран'}), 400
        
        # Проверяем расширение файла
        if not file.filename.lower().endswith('.docx'):
            return jsonify({'error': 'Поддерживаются только файлы с расширением .docx'}), 400
        
        # Читаем содержимое файла
        file_content = file.read()
        
        if not file_content:
            return jsonify({'error': 'Файл пустой'}), 400
        
        logger.info(f"Обработка файла: {file.filename}, размер: {len(file_content)} байт")
        
        # Обрабатываем файл
        text_result = process_docx_file(file_content)
        
        # Возвращаем результат
        response_data = {
            'success': True,
            'filename': file.filename,
            'text': text_result,
            'text_length': len(text_result)
        }
        
        logger.info(f"Файл {file.filename} успешно обработан. Длина текста: {len(text_result)} символов")
        
        return jsonify(response_data)
        
    except Exception as e:
        error_message = str(e)
        logger.error(f"Ошибка при обработке файла: {error_message}")
        
        return jsonify({
            'success': False,
            'error': error_message
        }), 500


@app.route('/convert-with-tables', methods=['POST'])
def convert_docx_with_tables():
    """
    Обработчик для конвертации DOCX файла в текст с таблицами
    
    Returns:
        JSON с результатом или ошибкой
    """
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'Файл не найден в запросе. Используйте поле "file"'}), 400
            
        file = request.files['file']
        
        if file.filename == '':
            return jsonify({'error': 'Файл не выбран'}), 400
        
        if not file.filename.lower().endswith('.docx'):
            return jsonify({'error': 'Поддерживаются только файлы с расширением .docx'}), 400
        
        file_content = file.read()
        
        if not file_content:
            return jsonify({'error': 'Файл пустой'}), 400
        
        logger.info(f"Обработка файла с таблицами: {file.filename}, размер: {len(file_content)} байт")
        
        # Обрабатываем файл с таблицами
        text_result = process_docx_with_tables(file_content)
        
        response_data = {
            'success': True,
            'filename': file.filename,
            'text': text_result,
            'text_length': len(text_result)
        }
        
        logger.info(f"Файл {file.filename} с таблицами успешно обработан. Длина текста: {len(text_result)} символов")
        
        return jsonify(response_data)
        
    except Exception as e:
        error_message = str(e)
        logger.error(f"Ошибка при обработке файла с таблицами: {error_message}")
        
        return jsonify({
            'success': False,
            'error': error_message
        }), 500


@app.route('/health', methods=['GET'])
def health_check():
    """
    Проверка работоспособности сервиса
    """
    return jsonify({
        'status': 'healthy',
        'service': 'DOCX to Text Converter'
    })


@app.errorhandler(413)
def too_large(e):
    """
    Обработчик ошибки превышения размера файла
    """
    return jsonify({
        'success': False,
        'error': 'Файл слишком большой. Максимальный размер: 50MB'
    }), 413


@app.errorhandler(400)
def bad_request(e):
    """
    Обработчик ошибок неверного запроса
    """
    return jsonify({
        'success': False,
        'error': 'Неверный запрос'
    }), 400


@app.errorhandler(500)
def internal_error(e):
    """
    Обработчик внутренних ошибок сервера
    """
    return jsonify({
        'success': False,
        'error': 'Внутренняя ошибка сервера'
    }), 500


def create_app():
    """
    Фабрика приложения
    
    Returns:
        Flask app
    """
    return app


def run_server(port=8000, host='0.0.0.0', debug=False):
    """
    Запуск HTTP сервера
    
    Args:
        port (int): Порт для запуска сервера
        host (str): Хост для привязки сервера
        debug (bool): Режим отладки
    """
    logger.info(f"Запуск сервера на {host}:{port}")
    app.run(host=host, port=port, debug=debug)
