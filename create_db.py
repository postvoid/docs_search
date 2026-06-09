#!/usr/bin/env python3
"""
СОЗДАНИЕ БАЗЫ ДАННЫХ С ПОЛНОТЕКСТОВЫМ ПОИСКОМ
Импорт данных из CSV файла (созданного краулером) в SQLite с FTS5
"""

import sqlite3
import os
import csv

DB_PATH = "output/search.db"
CSV_PATH = "output/documents_index.csv"


def create_database():
    """Создание БД и полнотекстового индекса"""
    if os.path.exists(DB_PATH):
        os.remove(DB_PATH)
        print("🗑 Удалена старая БД")

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    # FTS5 виртуальная таблица для полнотекстового поиска
    cursor.execute("""
        CREATE VIRTUAL TABLE documents 
        USING fts5(
            file_path,
            file_name, 
            file_type,
            content,
            tokenize='unicode61'
        );
    """)

    conn.commit()
    conn.close()
    print("✅ База данных создана с поддержкой FTS5")


def read_csv_content(file_path):
    """Чтение полного содержимого файла из CSV (если есть) или из оригинального файла"""
    # Пытаемся найти оригинальный файл и прочитать его содержимое
    if os.path.exists(file_path):
        try:
            ext = os.path.splitext(file_path)[1].lower()
            if ext in [".docx", ".doc", ".xlsx", ".xls", ".pdf", ".txt"]:
                with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                    return f.read()[:10000]  # Ограничиваем размер
        except:
            pass
    return ""


def import_from_csv():
    """Импорт данных из CSV в БД"""
    if not os.path.exists(CSV_PATH):
        print(f"❌ Файл {CSV_PATH} не найден!")
        print("Сначала запустите: python crawler.py")
        return False

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    # Читаем CSV файл
    print(f"📄 Чтение CSV файла: {CSV_PATH}")

    with open(CSV_PATH, "r", encoding="utf-8") as csvfile:
        reader = csv.DictReader(csvfile)
        rows = list(reader)

    print(f"📊 Загружено записей из CSV: {len(rows)}")

    # Для каждой записи в CSV пытаемся получить полное содержимое
    imported = 0
    for row in rows:
        file_path = row["file_path"]
        file_name = row["file_name"]
        file_type = row["file_type"]

        # Пытаемся прочитать оригинальный файл для получения полного текста
        content = read_csv_content(file_path)

        # Если не удалось прочитать файл, используем preview из CSV
        if not content:
            content = row.get("content_preview", "")
            print(f"  ⚠ Использую preview для: {file_name}")

        if content:
            try:
                cursor.execute(
                    """
                    INSERT INTO documents (file_path, file_name, file_type, content)
                    VALUES (?, ?, ?, ?)
                """,
                    (file_path, file_name, file_type, content),
                )
                imported += 1

                if imported % 10 == 0:
                    print(f"  Импортировано {imported} документов...")

            except Exception as e:
                print(f"  ✗ Ошибка импорта {file_name}: {e}")
        else:
            print(f"  ⚠ Нет содержимого для: {file_name}")

    conn.commit()

    # Проверка
    cursor.execute("SELECT COUNT(*) FROM documents")
    count = cursor.fetchone()[0]

    conn.close()

    print(f"\n✅ Импортировано {imported} документов из {len(rows)}")
    print(f"📊 Всего записей в БД: {count}")

    return imported > 0


def test_search():
    """Тест полнотекстового поиска"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    # Проверяем, есть ли данные
    cursor.execute("SELECT COUNT(*) FROM documents")
    count = cursor.fetchone()[0]

    if count == 0:
        print("\n⚠ Нет данных для поиска!")
        conn.close()
        return

    print("\n" + "=" * 70)
    print("ТЕСТ ПОЛНОТЕКСТОВОГО ПОИСКА")
    print("=" * 70)

    test_queries = ["финансовый", "договор", "техническое", "трудовой"]

    for query in test_queries:
        cursor.execute(
            """
            SELECT file_name, file_type
            FROM documents 
            WHERE documents MATCH ?
            LIMIT 5
        """,
            (query,),
        )

        results = cursor.fetchall()
        print(f"\n🔍 '{query}': найдено {len(results)} документов")
        for file_name, file_type in results[:3]:
            print(f"   - {file_name} ({file_type})")

    conn.close()


def show_csv_sample():
    """Показать пример данных из CSV"""
    if not os.path.exists(CSV_PATH):
        return

    print("\n" + "=" * 70)
    print("ПРИМЕР ДАННЫХ ИЗ CSV ФАЙЛА")
    print("=" * 70)

    with open(CSV_PATH, "r", encoding="utf-8") as csvfile:
        reader = csv.DictReader(csvfile)
        for i, row in enumerate(reader):
            if i >= 5:
                break
            print(f"\n{i+1}. {row['file_name']} [{row['file_type']}]")
            print(f"   Путь: {row['file_path']}")
            print(f"   Размер: {row['file_size']} байт")
            preview = row["content_preview"][:100]
            print(f"   Превью: {preview}...")


if __name__ == "__main__":
    print("=" * 70)
    print("СОЗДАНИЕ БАЗЫ ДАННЫХ С ПОЛНОТЕКСТОВЫМ ПОИСКОМ")
    print("Импорт из CSV файла (созданного краулером)")
    print("=" * 70)

    # Показываем пример CSV
    show_csv_sample()

    # Создаём БД
    create_database()

    # Импортируем из CSV
    if import_from_csv():
        test_search()

        print("\n" + "=" * 70)
        print("✅ ГОТОВО! Запустите поиск: python search.py")
        print("=" * 70)
    else:
        print("\n❌ ОШИБКА: Не удалось импортировать данные")
        print("Проверьте, что CSV файл существует и содержит данные")
