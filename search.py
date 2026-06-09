#!/usr/bin/env python3
"""
ПОИСКОВАЯ СИСТЕМА
Полнотекстовый поиск по документам
"""

import sqlite3
import sys
import os

DB_PATH = "output/search.db"


def search(query):
    if not os.path.exists(DB_PATH):
        print(f"❌ БД не найдена")
        return

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    cursor.execute("SELECT COUNT(*) FROM documents")
    total = cursor.fetchone()[0]

    print(f"\n🔍 ПОИСК: '{query}'")
    print("=" * 70)
    print(f"📊 Индексировано документов: {total}\n")

    cursor.execute(
        """
        SELECT file_name, file_type, highlight(documents, 0, '«', '»') as snippet
        FROM documents WHERE content MATCH ? ORDER BY rank LIMIT 30
    """,
        (query,),
    )

    results = cursor.fetchall()

    if not results:
        print("❌ НИЧЕГО НЕ НАЙДЕНО")
        conn.close()
        return

    print(f"✅ НАЙДЕНО: {len(results)} ДОКУМЕНТОВ\n")
    print("-" * 70)

    for i, (name, typ, snippet) in enumerate(results, 1):
        print(f"{i}. {name} [{typ}]")
        if snippet:
            print(f"   → {snippet[:150]}...")
        print()

    conn.close()


def interactive():
    print("=" * 70)
    print("🔎 ПОИСКОВАЯ СИСТЕМА")
    print("=" * 70)

    if os.path.exists(DB_PATH):
        conn = sqlite3.connect(DB_PATH)
        cur = conn.cursor()
        cur.execute("SELECT COUNT(*) FROM documents")
        print(f"📁 Документов в индексе: {cur.fetchone()[0]}")
        conn.close()

    print("\nКоманды: /exit - выход, /stats - статистика")
    print("=" * 70)

    while True:
        try:
            query = input("\n🔍 Введите запрос > ").strip()
            if query == "/exit":
                break
            elif query == "/stats":
                conn = sqlite3.connect(DB_PATH)
                cur = conn.cursor()
                cur.execute(
                    "SELECT file_type, COUNT(*) FROM documents GROUP BY file_type"
                )
                for t, c in cur.fetchall():
                    print(f"  {t}: {c}")
                conn.close()
            elif query:
                search(query)
        except KeyboardInterrupt:
            print("\nДо свидания!")
            break


if __name__ == "__main__":
    if len(sys.argv) > 1:
        search(" ".join(sys.argv[1:]))
    else:
        interactive()
