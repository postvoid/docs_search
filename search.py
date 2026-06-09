#!/usr/bin/env python3
"""
ПОИСКОВАЯ СИСТЕМА
Полнотекстовый поиск по документам в SQLite FTS5
"""

import sqlite3
import sys
import os

DB_PATH = "output/search.db"


def search(query):
    """Полнотекстовый поиск"""
    if not os.path.exists(DB_PATH):
        print(f"❌ База данных не найдена: {DB_PATH}")
        print("Сначала запустите: python create_db.py")
        return

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    # Проверяем, есть ли данные
    cursor.execute("SELECT COUNT(*) FROM documents")
    count = cursor.fetchone()[0]

    if count == 0:
        print("⚠ База данных пуста")
        conn.close()
        return

    print(f"\n🔍 ПОИСК: '{query}'")
    print("=" * 70)
    print(f"📊 Индексировано документов: {count}\n")

    # Выполняем поиск
    cursor.execute(
        """
        SELECT file_name, file_type, 
               highlight(documents, 3, '<b>', '</b>') as snippet
        FROM documents 
        WHERE documents MATCH ?
        ORDER BY rank
        LIMIT 20
    """,
        (query,),
    )

    results = cursor.fetchall()

    if not results:
        # Пробуем поиск по отдельным словам
        words = query.split()
        if len(words) > 1:
            fts_query = " OR ".join(words)
            cursor.execute(
                """
                SELECT file_name, file_type,
                       highlight(documents, 3, '<b>', '</b>') as snippet
                FROM documents 
                WHERE documents MATCH ?
                ORDER BY rank
                LIMIT 20
            """,
                (fts_query,),
            )
            results = cursor.fetchall()

    if not results:
        print("❌ НИЧЕГО НЕ НАЙДЕНО")
        print("\n💡 ПРИМЕРЫ ЗАПРОСОВ:")
        print("   • финансовый")
        print("   • договор")
        print("   • техническое задание")
        print("   • трудовой договор")
        print("   • бюджет")
        conn.close()
        return

    print(f"✅ НАЙДЕНО: {len(results)} ДОКУМЕНТОВ\n")
    print("-" * 70)

    for i, (file_name, file_type, snippet) in enumerate(results, 1):
        print(f"{i}. {file_name} [{file_type}]")
        if snippet:
            # Очищаем от HTML тегов для отображения
            clean_snippet = snippet.replace("<b>", "**").replace("</b>", "**")
            print(f"   → {clean_snippet[:200]}...")
        print()

    conn.close()


def interactive():
    """Интерактивный режим"""
    print("=" * 70)
    print("🔎 ПОИСКОВАЯ СИСТЕМА ПО ДОКУМЕНТАМ")
    print("=" * 70)

    if os.path.exists(DB_PATH):
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM documents")
        count = cursor.fetchone()[0]
        conn.close()
        print(f"📁 Документов в индексе: {count}")
    else:
        print("⚠ База данных не создана")
        print("Сначала запустите: python create_db.py")
        return

    print("\nКоманды:")
    print("  /exit - выход")
    print("  /help - помощь")
    print("  /stats - статистика")
    print("=" * 70)

    while True:
        try:
            query = input("\n🔍 Введите запрос > ").strip()

            if query.lower() == "/exit":
                print("До свидания!")
                break
            elif query.lower() == "/help":
                print("\n📖 ПРИМЕРЫ ЗАПРОСОВ:")
                print("   финансовый - поиск по слову")
                print("   договор соглашение - поиск по нескольким словам")
                print("   техническое задание - поиск фразы")
                print("\n🏷️  КЛЮЧЕВЫЕ СЛОВА В ДОКУМЕНТАХ:")
                print("   finance: финансовый, бюджет, налог, инвестиционный")
                print("   legal: договор, соглашение, иск, претензия, лицензионное")
                print("   tech: техническое, руководство, архитектура, инструкция")
                print("   hr: трудовой, приказ, штатное, должностная")
            elif query.lower() == "/stats":
                conn = sqlite3.connect(DB_PATH)
                cursor = conn.cursor()
                cursor.execute(
                    "SELECT file_type, COUNT(*) FROM documents GROUP BY file_type"
                )
                stats = cursor.fetchall()
                print("\n📊 СТАТИСТИКА:")
                for file_type, count in stats:
                    print(f"   {file_type}: {count} файлов")
                conn.close()
            elif query:
                search(query)

        except KeyboardInterrupt:
            print("\n\nДо свидания!")
            break
        except Exception as e:
            print(f"❌ Ошибка: {e}")


if __name__ == "__main__":
    if len(sys.argv) > 1:
        search(" ".join(sys.argv[1:]))
    else:
        interactive()
