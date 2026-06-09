#!/usr/bin/env python3
"""
СОЗДАНИЕ БАЗЫ ДАННЫХ С ПОЛНОТЕКСТОВЫМ ПОИСКОМ
Импорт данных из CSV в SQLite с FTS5
"""

import sqlite3
import pandas as pd
import pickle
import os

DB_PATH = "output/search.db"
CSV_PATH = "output/documents_index.csv"
TEXTS_PATH = "output/full_texts.pkl"


def create_database():
    if os.path.exists(DB_PATH):
        os.remove(DB_PATH)
        print("🗑 Удалена старая БД")

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("""
        CREATE VIRTUAL TABLE documents 
        USING fts5(content, file_name, file_type, file_path, tokenize='unicode61')
    """)
    conn.commit()
    conn.close()
    print("✅ База данных создана")


def import_data():
    if not os.path.exists(CSV_PATH):
        print(f"❌ Файл {CSV_PATH} не найден!")
        return False

    if not os.path.exists(TEXTS_PATH):
        print(f"❌ Файл {TEXTS_PATH} не найден!")
        return False

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    df = pd.read_csv(CSV_PATH)
    with open(TEXTS_PATH, "rb") as f:
        full_texts = pickle.load(f)

    texts_dict = {item["file_path"]: item["content"] for item in full_texts}

    imported = 0
    for _, row in df.iterrows():
        file_path = row["file_path"]
        content = texts_dict.get(file_path, "")
        if content:
            cursor.execute(
                """
                INSERT INTO documents (content, file_name, file_type, file_path)
                VALUES (?, ?, ?, ?)
            """,
                (content, row["file_name"], row["file_type"], file_path),
            )
            imported += 1

    conn.commit()
    cursor.execute("SELECT COUNT(*) FROM documents")
    count = cursor.fetchone()[0]
    conn.close()

    print(f"✅ Импортировано {imported} документов в БД")
    return True


def test_search():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) FROM documents")
    count = cursor.fetchone()[0]

    print("\n" + "=" * 70)
    print("ТЕСТ ПОИСКА")
    print("=" * 70)

    for query in ["финансовый", "договор", "техническое", "трудовой"]:
        cursor.execute(
            "SELECT file_name, file_type FROM documents WHERE content MATCH ? LIMIT 5",
            (query,),
        )
        results = cursor.fetchall()
        print(f"\n🔍 '{query}': {len(results)} результатов")
        for name, typ in results[:3]:
            print(f"   - {name} ({typ})")
    conn.close()


if __name__ == "__main__":
    print("=" * 70)
    print("СОЗДАНИЕ БД С ПОЛНОТЕКСТОВЫМ ПОИСКОМ")
    print("=" * 70)
    create_database()
    if import_data():
        test_search()
        print("\n✅ ГОТОВО! Запустите: python search.py")
