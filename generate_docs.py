#!/usr/bin/env python3
"""
ГЕНЕРАТОР ХРАНИЛИЩА ФАЙЛОВ
Создаёт: .doc, .docx, .xls, .xlsx, .pdf, .zip, .rar, .7z
"""

import os
import random
import zipfile
import subprocess
from datetime import datetime
from docx import Document
from openpyxl import Workbook
import xlwt

# Пробуем импортировать FPDF
try:
    from fpdf import FPDF

    FPDF_AVAILABLE = True
except ImportError:
    FPDF_AVAILABLE = False
    print("⚠ fpdf не установлен. PDF созданы не будут.")

# Пробуем импортировать rarfile для создания RAR
try:
    import rarfile

    RARFILE_AVAILABLE = True
except ImportError:
    RARFILE_AVAILABLE = False
    print("⚠ rarfile не установлен. RAR созданы не будут.")

# ==============================
# НАСТРОЙКА ДИРЕКТОРИЙ
# ==============================
STORAGE_DIR = "storage"
DOCS_DIR = os.path.join(STORAGE_DIR, "docs")
ARCHIVES_DIR = os.path.join(STORAGE_DIR, "archives")

os.makedirs(DOCS_DIR, exist_ok=True)
os.makedirs(ARCHIVES_DIR, exist_ok=True)

# ==============================
# ТЕМАТИКИ ДОКУМЕНТОВ
# ==============================
DOCUMENTS = [
    (
        "Финансовый отчёт",
        "finance",
        "Анализ финансовых показателей компании за год. Выручка выросла на 25%.",
    ),
    (
        "Договор оказания услуг",
        "legal",
        "Договор между исполнителем и заказчиком. Стоимость услуг: 100 000 рублей.",
    ),
    (
        "Техническое задание",
        "tech",
        "Требования к разработке программного обеспечения. Использовать Python.",
    ),
    (
        "Трудовой договор",
        "hr",
        "Договор между работником и работодателем. Оклад: 50 000 рублей.",
    ),
    (
        "Налоговая декларация",
        "finance",
        "Отчётность по налогам за квартал. Сумма налога: 25 000 рублей.",
    ),
    (
        "Лицензионное соглашение",
        "legal",
        "Правила использования программного продукта. Лицензия: бессрочная.",
    ),
    (
        "Руководство пользователя",
        "tech",
        "Инструкция по работе с системой. Для начинающих пользователей.",
    ),
    (
        "Приказ о приёме",
        "hr",
        "Приказ о зачислении сотрудника в штат. Должность: разработчик.",
    ),
    (
        "Бюджет проекта",
        "finance",
        "Планирование расходов на проект. Бюджет: 1 000 000 рублей.",
    ),
    (
        "Акт приёма-передачи",
        "legal",
        "Акт сдачи-приёмки выполненных работ. Работы выполнены полностью.",
    ),
]


# ==============================
# ПОИСК ШРИФТА ДЛЯ PDF
# ==============================
def find_font():
    """Поиск шрифта с поддержкой кириллицы"""
    possible_paths = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        "/usr/share/fonts/TTF/DejaVuSans.ttf",
        "/usr/share/fonts/dejavu/DejaVuSans.ttf",
        "/System/Library/Fonts/Helvetica.ttc",
    ]

    for path in possible_paths:
        if os.path.exists(path):
            return path
    return None


# ==============================
# ГЕНЕРАЦИЯ .DOCX
# ==============================
def generate_docx(filename, title, category, content):
    doc = Document()
    doc.add_heading(title, level=1)
    doc.add_paragraph(f"Категория: {category}")
    doc.add_paragraph(f"Дата создания: {datetime.now()}")
    doc.add_paragraph(f"Ключевые слова: {category}, {title}, поиск, индексация, тест")
    doc.add_paragraph(f"Содержание: {content}. " * 30)
    doc.add_paragraph("Дополнительный текст для полнотекстового поиска. " * 20)
    doc.save(filename)


# ==============================
# ГЕНЕРАЦИЯ .DOC
# ==============================
def generate_doc(filename, title, category, content):
    with open(filename, "w", encoding="utf-8") as f:
        f.write(f"{title}\n")
        f.write(f"{'='*60}\n")
        f.write(f"Категория: {category}\n")
        f.write(f"Дата создания: {datetime.now()}\n")
        f.write(f"Ключевые слова: {category}, {title}, поиск, индексация, тест\n")
        f.write(f"\nСодержание:\n")
        f.write(f"{content}. " * 50)


# ==============================
# ГЕНЕРАЦИЯ .XLSX
# ==============================
def generate_xlsx(filename, title, category, content):
    wb = Workbook()
    ws = wb.active
    ws.title = "Документ"

    headers = ["Название", "Категория", "Содержание", "Дата", "Ключевые слова"]
    for col, header in enumerate(headers, 1):
        ws.cell(row=1, column=col, value=header)

    ws.cell(row=2, column=1, value=title)
    ws.cell(row=2, column=2, value=category)
    ws.cell(row=2, column=3, value=content)
    ws.cell(row=2, column=4, value=str(datetime.now()))
    ws.cell(row=2, column=5, value=f"{category}, поиск, тест")

    for i in range(3, 15):
        ws.cell(row=i, column=1, value=f"Запись {i-2}")
        ws.cell(row=i, column=2, value=f"Значение {random.randint(1, 100)}")
        ws.cell(row=i, column=3, value=f"Поисковый текст: {title}")

    wb.save(filename)


# ==============================
# ГЕНЕРАЦИЯ .XLS
# ==============================
def generate_xls(filename, title, category, content):
    wb = xlwt.Workbook(encoding="utf-8")
    ws = wb.add_sheet("Документ")

    headers = ["Название", "Категория", "Содержание", "Дата", "Ключевые слова"]
    for col, header in enumerate(headers):
        ws.write(0, col, header)

    ws.write(1, 0, title)
    ws.write(1, 1, category)
    ws.write(1, 2, content)
    ws.write(1, 3, str(datetime.now()))
    ws.write(1, 4, f"{category}, поиск, тест")

    for i in range(2, 12):
        ws.write(i, 0, f"Запись {i-1}")
        ws.write(i, 1, f"Значение {random.randint(1, 100)}")
        ws.write(i, 2, f"Поисковый текст: {title}")

    wb.save(filename)


# ==============================
# ГЕНЕРАЦИЯ .PDF
# ==============================
def generate_pdf(filename, title, category, content, font_path):
    try:
        pdf = FPDF()
        pdf.add_page()
        pdf.add_font("Russian", "", font_path, uni=True)
        pdf.set_font("Russian", "", 14)
        pdf.cell(200, 10, text=title, new_x="LMARGIN", new_y="NEXT", align="C")
        pdf.ln(10)
        pdf.set_font("Russian", "", 12)
        pdf.cell(200, 8, text=f"Категория: {category}", new_x="LMARGIN", new_y="NEXT")
        pdf.cell(200, 8, text=f"Дата: {datetime.now()}", new_x="LMARGIN", new_y="NEXT")
        pdf.cell(
            200,
            8,
            text=f"Ключевые слова: {category}, {title}, поиск",
            new_x="LMARGIN",
            new_y="NEXT",
        )
        pdf.ln(10)
        pdf.multi_cell(0, 6, text=f"Содержание: {content}. " * 30)
        pdf.output(filename)
        return True
    except Exception as e:
        return False


# ==============================
# СОЗДАНИЕ RAR АРХИВА (через подпроцесс, если есть rar)
# ==============================
def create_rar_archive(files, output_path, files_count=5):
    """Создание RAR архива с помощью системной утилиты rar или альтернатив"""

    # Проверяем, есть ли системная утилита rar
    try:
        subprocess.run(["rar", "--version"], capture_output=True, check=True)
        rar_cmd = "rar"
    except (subprocess.CalledProcessError, FileNotFoundError):
        rar_cmd = None

    if rar_cmd:
        # Используем системную утилиту rar
        files_str = " ".join(f'"{f}"' for f in files if os.path.exists(f))
        if files_str:
            result = subprocess.run(
                f"{rar_cmd} a -ep1 -m0 {output_path} {files_str}",
                shell=True,
                capture_output=True,
            )
            return os.path.exists(output_path)

    # Если нет rar, но есть unrar (только для чтения), пропускаем
    # Альтернатива: используем zip как fallback
    print(f"    rar не найден, используем zip вместо rar")
    zip_path = output_path.replace(".rar", ".rar.zip")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in files:
            if os.path.exists(f):
                zf.write(f, os.path.basename(f))
    if os.path.exists(zip_path):
        os.rename(zip_path, output_path)
        return True

    return False


# ==============================
# ГЕНЕРАЦИЯ АРХИВОВ
# ==============================
def create_archives(all_files):
    """Создание архивов всех форматов"""

    # 1. ZIP архив
    zip_path = os.path.join(ARCHIVES_DIR, "documents.zip")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in all_files[:5]:
            if os.path.exists(f):
                zf.write(f, os.path.basename(f))
    print(f"  ✓ documents.zip (5 файлов)")

    # 2. RAR архив
    rar_path = os.path.join(ARCHIVES_DIR, "documents.rar")
    rar_files = [f for f in all_files[5:10] if os.path.exists(f)]
    if rar_files:
        if create_rar_archive(rar_files, rar_path, 5):
            print(f"  ✓ documents.rar (5 файлов)")
        else:
            print(f"  ✗ documents.rar не создан")
    else:
        print(f"  ⚠ Нет файлов для RAR")

    # 3. 7z архив
    sz_available = False
    try:
        subprocess.run(["7z", "--help"], capture_output=True, check=True)
        sz_available = True
    except (subprocess.CalledProcessError, FileNotFoundError):
        pass

    if sz_available:
        sz_path = os.path.join(ARCHIVES_DIR, "documents.7z")
        sz_files = [f for f in all_files[10:15] if os.path.exists(f)]
        if sz_files:
            files_str = " ".join(f'"{f}"' for f in sz_files)
            subprocess.run(
                f"7z a -tzip {sz_path} {files_str}", shell=True, capture_output=True
            )
            if os.path.exists(sz_path):
                print(f"  ✓ documents.7z (5 файлов)")
            else:
                print(f"  ✗ documents.7z не создан")
        else:
            print(f"  ⚠ Нет файлов для 7z")
    else:
        print(f"  ⚠ 7z архив не создан (установите: sudo pacman -S p7zip)")


# ==============================
# ОСНОВНАЯ ФУНКЦИЯ
# ==============================
def main():
    print("=" * 70)
    print("ГЕНЕРАТОР ХРАНИЛИЩА ФАЙЛОВ")
    print("Форматы: .doc, .docx, .xls, .xlsx, .pdf, .zip, .rar, .7z")
    print("=" * 70)

    # Поиск шрифта для PDF
    print("\n0. Подготовка PDF...")
    font_path = find_font()
    if font_path:
        print(f"  ✓ Шрифт найден: {font_path}")
    else:
        print(f"  ✗ Шрифт не найден! PDF созданы не будут.")

    all_files = []

    # Генерация документов
    print("\n1. Генерация документов:")
    print("-" * 50)

    pdf_ok = 0
    for i, (title, category, content) in enumerate(DOCUMENTS, 1):
        base = f"doc_{i}_{category}"

        # DOCX
        path = os.path.join(DOCS_DIR, f"{base}.docx")
        generate_docx(path, title, category, content)
        all_files.append(path)
        print(f"  ✓ {base}.docx")

        # DOC
        path = os.path.join(DOCS_DIR, f"{base}.doc")
        generate_doc(path, title, category, content)
        all_files.append(path)
        print(f"  ✓ {base}.doc")

        # XLSX
        path = os.path.join(DOCS_DIR, f"{base}.xlsx")
        generate_xlsx(path, title, category, content)
        all_files.append(path)
        print(f"  ✓ {base}.xlsx")

        # XLS
        path = os.path.join(DOCS_DIR, f"{base}.xls")
        generate_xls(path, title, category, content)
        all_files.append(path)
        print(f"  ✓ {base}.xls")

        # PDF
        if font_path:
            path = os.path.join(DOCS_DIR, f"{base}.pdf")
            if generate_pdf(path, title, category, content, font_path):
                all_files.append(path)
                pdf_ok += 1
                print(f"  ✓ {base}.pdf")
            else:
                print(f"  ✗ {base}.pdf - ошибка")
        else:
            print(f"  ⚠ {base}.pdf - пропущен")

    print("-" * 50)
    print(f"  Всего файлов: {len(all_files)}")
    if font_path:
        print(f"  PDF успешно: {pdf_ok}/{len(DOCUMENTS)}")

    # Генерация архивов
    print("\n2. Создание архивов с вложенными файлами:")
    print("-" * 50)
    create_archives(all_files)

    # Статистика
    print("\n" + "=" * 70)
    print("СТАТИСТИКА ХРАНИЛИЩА")
    print("=" * 70)

    doc_count = len(
        [f for f in os.listdir(DOCS_DIR) if os.path.isfile(os.path.join(DOCS_DIR, f))]
    )
    arch_count = len(
        [
            f
            for f in os.listdir(ARCHIVES_DIR)
            if os.path.isfile(os.path.join(ARCHIVES_DIR, f))
        ]
    )

    total_size = 0
    for root, dirs, files in os.walk(STORAGE_DIR):
        for file in files:
            total_size += os.path.getsize(os.path.join(root, file))

    print(f"\n📁 Хранилище: {STORAGE_DIR}")
    print(f"   - Документы: {doc_count} файлов")
    print(f"   - Архивы: {arch_count} файлов")
    print(f"   - Общий размер: {total_size / 1024:.1f} KB")

    print("\n" + "=" * 70)
    print("✅ ГЕНЕРАЦИЯ ЗАВЕРШЕНА")
    print("=" * 70)
    print("\n🎯 ДАЛЕЕ ЗАПУСТИТЕ: python crawler.py")


if __name__ == "__main__":
    main()


def main():
    print("=" * 70)
    print("ГЕНЕРАТОР ХРАНИЛИЩА ФАЙЛОВ")
    print("Форматы: .doc, .docx, .xls, .xlsx, .pdf, .zip, .rar, .7z")
    print("=" * 70)

    font_path = find_font()
    if font_path:
        print(f"\n✓ Шрифт найден: {font_path}")
    else:
        print("\n⚠ Шрифт не найден, PDF не созданы")

    all_files = []
    print("\n1. Генерация документов:")
    print("-" * 50)

    for i, (title, category, content) in enumerate(DOCUMENTS, 1):
        base = f"doc_{i}_{category}"

        path = os.path.join(DOCS_DIR, f"{base}.docx")
        generate_docx(path, title, category, content)
        all_files.append(path)
        print(f"  ✓ {base}.docx")

        path = os.path.join(DOCS_DIR, f"{base}.doc")
        generate_doc(path, title, category, content)
        all_files.append(path)
        print(f"  ✓ {base}.doc")

        path = os.path.join(DOCS_DIR, f"{base}.xlsx")
        generate_xlsx(path, title, category, content)
        all_files.append(path)
        print(f"  ✓ {base}.xlsx")

        path = os.path.join(DOCS_DIR, f"{base}.xls")
        generate_xls(path, title, category, content)
        all_files.append(path)
        print(f"  ✓ {base}.xls")

        if font_path:
            path = os.path.join(DOCS_DIR, f"{base}.pdf")
            if generate_pdf(path, title, category, content, font_path):
                all_files.append(path)
                print(f"  ✓ {base}.pdf")
            else:
                print(f"  ✗ {base}.pdf")

    print("-" * 50)
    print(f"  Всего файлов: {len(all_files)}")

    print("\n2. Создание архивов:")
    print("-" * 50)
    create_archives(all_files)

    print("\n" + "=" * 70)
    print("✅ ГЕНЕРАЦИЯ ЗАВЕРШЕНА")
    print("=" * 70)


if __name__ == "__main__":
    main()
