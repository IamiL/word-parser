#!/usr/bin/env python3
"""
Точка входа в программу для конвертации DOCX файлов в текст
"""

import sys
import argparse
import logging
from server import run_server


def parse_arguments():
    """
    Парсинг аргументов командной строки
    
    Returns:
        argparse.Namespace: Объект с аргументами
    """
    parser = argparse.ArgumentParser(
        description='HTTP сервер для конвертации DOCX файлов в текст',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Примеры использования:
  python main.py --port 8080
  python main.py -p 5000 --host 127.0.0.1
  python main.py --port 8000 --debug

После запуска сервера используйте:
  curl -X POST -F 'file=@document.docx' http://localhost:8000/convert
        """
    )
    
    parser.add_argument(
        '-p', '--port',
        type=int,
        default=8000,
        help='Порт для запуска HTTP сервера (по умолчанию: 8000)'
    )
    
    parser.add_argument(
        '--host',
        type=str,
        default='0.0.0.0',
        help='Хост для привязки сервера (по умолчанию: 0.0.0.0)'
    )
    
    parser.add_argument(
        '--debug',
        action='store_true',
        help='Запуск в режиме отладки'
    )
    
    parser.add_argument(
        '--log-level',
        choices=['DEBUG', 'INFO', 'WARNING', 'ERROR'],
        default='INFO',
        help='Уровень логирования (по умолчанию: INFO)'
    )
    
    return parser.parse_args()


def setup_logging(level):
    """
    Настройка системы логирования
    
    Args:
        level (str): Уровень логирования
    """
    numeric_level = getattr(logging, level.upper(), None)
    if not isinstance(numeric_level, int):
        raise ValueError(f'Неверный уровень логирования: {level}')
    
    logging.basicConfig(
        level=numeric_level,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )


def validate_port(port):
    """
    Валидация номера порта
    
    Args:
        port (int): Номер порта
        
    Returns:
        bool: True если порт валиден
        
    Raises:
        ValueError: При невалидном порте
    """
    if not (1 <= port <= 65535):
        raise ValueError(f"Порт должен быть в диапазоне 1-65535, получен: {port}")
    
    if port < 1024:
        logging.warning(f"Используется привилегированный порт {port}. "
                       f"Может потребоваться запуск с правами администратора.")
    
    return True


def print_startup_info(host, port, debug):
    """
    Вывод информации о запуске сервера
    
    Args:
        host (str): Хост
        port (int): Порт
        debug (bool): Режим отладки
    """
    print("=" * 60)
    print("  DOCX to Text Converter Server")
    print("=" * 60)
    print(f"Сервер запущен на: http://{host}:{port}")
    print(f"Режим отладки: {'Включен' if debug else 'Выключен'}")
    print("")
    print("Доступные эндпоинты:")
    print(f"  GET  http://{host}:{port}/           - Информация об API")
    print(f"  GET  http://{host}:{port}/health     - Проверка работоспособности")
    print(f"  POST http://{host}:{port}/convert    - Конвертация DOCX в текст")
    print(f"  POST http://{host}:{port}/convert-with-tables - С таблицами")
    print("")
    print("Пример использования:")
    print(f"  curl -X POST -F 'file=@document.docx' http://{host}:{port}/convert")
    print("")
    print("Для остановки сервера нажмите Ctrl+C")
    print("=" * 60)


def main():
    """
    Главная функция программы
    """
    try:
        # Парсим аргументы командной строки
        args = parse_arguments()
        
        # Настраиваем логирование
        setup_logging(args.log_level)
        
        # Валидируем порт
        validate_port(args.port)
        
        # Выводим информацию о запуске
        print_startup_info(args.host, args.port, args.debug)
        
        # Запускаем сервер
        run_server(
            port=args.port,
            host=args.host,
            debug=args.debug
        )
        
    except KeyboardInterrupt:
        print("\n\nСервер остановлен пользователем")
        sys.exit(0)
        
    except ValueError as e:
        print(f"Ошибка валидации: {e}", file=sys.stderr)
        sys.exit(1)
        
    except Exception as e:
        print(f"Критическая ошибка: {e}", file=sys.stderr)
        logging.exception("Критическая ошибка при запуске")
        sys.exit(1)


if __name__ == '__main__':
    main()
