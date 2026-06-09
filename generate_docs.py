#!/usr/bin/env python3
"""
ГЕНЕРАТОР ХРАНИЛИЩА ФАЙЛОВ
Создаёт тестовые файлы форматов: .doc, .docx, .xls, .xlsx, .pdf
А также архивы: .zip, .rar, .7z с вложенными файлами
"""

import os
import random
import zipfile
import subprocess
from datetime import datetime
from docx import Document
from openpyxl import Workbook
import xlwt
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import urllib.request

# ==============================
# НАСТРОЙКА ДИРЕКТОРИЙ
# ==============================
STORAGE_DIR = "storage"
DOCS_DIR = os.path.join(STORAGE_DIR, "docs")
ARCHIVES_DIR = os.path.join(STORAGE_DIR, "archives")
FONTS_DIR = os.path.join(os.path.dirname(__file__), "fonts")

os.makedirs(DOCS_DIR, exist_ok=True)
os.makedirs(ARCHIVES_DIR, exist_ok=True)
os.makedirs(FONTS_DIR, exist_ok=True)


# ==============================
# ЗАГРУЗКА ШРИФТА ДЛЯ PDF
# ==============================
def download_font():
    """Скачивание шрифта DejaVu для поддержки русского языка в PDF"""
    font_path = os.path.join(FONTS_DIR, "DejaVuSans.ttf")

    if os.path.exists(font_path):
        return font_path

    print("  📥 Скачивание шрифта для PDF...")
    urls = [
        "https://github.com/dejavu-fonts/dejavu-fonts/raw/master/ttf/DejaVuSans.ttf",
        "https://raw.githubusercontent.com/dejavu-fonts/dejavu-fonts/master/ttf/DejaVuSans.ttf",
    ]

    for url in urls:
        try:
            urllib.request.urlretrieve(url, font_path)
            if os.path.exists(font_path) and os.path.getsize(font_path) > 100000:
                print(f"  ✓ Шрифт загружен")
                return font_path
        except:
            continue

    # Если не скачался, используем системный
    system_fonts = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/System/Library/Fonts/Helvetica.ttc",
        "/Windows/Fonts/arial.ttf",
    ]

    for sys_path in system_fonts:
        if os.path.exists(sys_path):
            print(f"  ✓ Используется системный шрифт")
            return sys_path

    raise Exception(
        "Шрифт не найден. Установите: sudo apt-get install fonts-dejavu-core"
    )


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
    (
        "Инвестиционный портфель",
        "finance",
        "Анализ доходности инвестиционного портфеля. Доходность: 15% годовых.",
    ),
    (
        "Претензионное письмо",
        "legal",
        "Письмо с требованием оплатить задолженность. Сумма: 50 000 рублей.",
    ),
    (
        "Архитектура системы",
        "tech",
        "Описание архитектуры программного комплекса. Микросервисная архитектура.",
    ),
    (
        "Штатное расписание",
        "hr",
        "Список должностей и окладов. Всего сотрудников: 50 человек.",
    ),
]


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
        f.write(f"\n\nПоисковый текст для проверки: {title} {category}")


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

    for i in range(3, 25):
        ws.cell(row=i, column=1, value=f"Дополнительная запись {i-2}")
        ws.cell(row=i, column=2, value=f"Значение {random.randint(1, 1000)}")
        ws.cell(
            row=i, column=3, value=f"Поисковый текст: {title} {category} {content[:50]}"
        )

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

    for i in range(2, 20):
        ws.write(i, 0, f"Запись {i-1}")
        ws.write(i, 1, f"Данные {random.randint(1, 1000)}")
        ws.write(i, 2, f"Поисковый текст: {title} {category}")

    wb.save(filename)


# ==============================
# ГЕНЕРАЦИЯ .PDF (с поддержкой русского языка)
# ==============================
def generate_pdf(filename, title, category, content, font_path):
    """Генерация PDF с поддержкой русского языка через reportlab"""
    # Регистрируем шрифт
    pdfmetrics.registerFont(TTFont("RussianFont", font_path))

    c = canvas.Canvas(filename, pagesize=A4)
    width, height = A4

    # Заголовок
    c.setFont("RussianFont", 18)
    c.drawString(50, height - 50, title)

    # Метаданные
    c.setFont("RussianFont", 12)
    c.drawString(50, height - 80, f"Категория: {category}")
    c.drawString(50, height - 100, f"Дата создания: {datetime.now()}")
    c.drawString(
        50,
        height - 120,
        f"Ключевые слова: {category}, {title}, поиск, индексация, тест",
    )

    # Содержание
    y = height - 160
    c.setFont("RussianFont", 10)
    text = f"Содержание: {content}. " * 30

    # Перенос текста
    lines = []
    current_line = ""
    for word in text.split():
        if c.stringWidth(current_line + " " + word, "RussianFont", 10) < width - 100:
            current_line += " " + word if current_line else word
        else:
            lines.append(current_line)
            current_line = word
    if current_line:
        lines.append(current_line)

    for line in lines:
        c.drawString(50, y, line)
        y -= 20
        if y < 50:
            c.showPage()
            c.setFont("RussianFont", 10)
            y = height - 50

    c.save()


# ==============================
# ГЕНЕРАЦИЯ АРХИВОВ
# ==============================
def create_archives(file_list):
    """Создание архивов с вложенными файлами"""
    archives_created = []

    # ZIP архив
    zip_path = os.path.join(ARCHIVES_DIR, "documents.zip")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in file_list[:8]:
            if os.path.exists(f):
                zf.write(f, os.path.basename(f))
    archives_created.append(zip_path)
    print(f"  ✓ documents.zip")

    # RAR архив (если установлен rar)
    rar_path = os.path.join(ARCHIVES_DIR, "documents.rar")
    try:
        files_str = " ".join(f'"{f}"' for f in file_list[8:16] if os.path.exists(f))
        if files_str:
            subprocess.run(
                f"rar a -ep1 {rar_path} {files_str}", shell=True, capture_output=True
            )
            if os.path.exists(rar_path):
                archives_created.append(rar_path)
                print(f"  ✓ documents.rar")
            else:
                print(f"  ⚠ documents.rar не создан (rar не установлен)")
    except:
        print(f"  ⚠ documents.rar не создан (rar не установлен)")

    # 7z архив (если установлен 7z)
    sz_path = os.path.join(ARCHIVES_DIR, "documents.7z")
    try:
        files_str = " ".join(f'"{f}"' for f in file_list[16:24] if os.path.exists(f))
        if files_str:
            subprocess.run(
                f"7z a {sz_path} {files_str}", shell=True, capture_output=True
            )
            if os.path.exists(sz_path):
                archives_created.append(sz_path)
                print(f"  ✓ documents.7z")
            else:
                print(f"  ⚠ documents.7z не создан (7z не установлен)")
    except:
        print(f"  ⚠ documents.7z не создан (7z не установлен)")

    return archives_created


# ==============================
# ОСНОВНАЯ ФУНКЦИЯ
# ==============================
def main():
    print("=" * 70)
    print("ГЕНЕРАТОР ХРАНИЛИЩА ФАЙЛОВ")
    print("Форматы: .doc, .docx, .xls, .xlsx, .pdf, .zip, .rar, .7z")
    print("=" * 70)

    # Загрузка шрифта для PDF
    print("\n0. Подготовка шрифтов для PDF...")
    try:
        font_path = download_font()
        print("  ✓ Шрифт готов\n")
    except Exception as e:
        print(f"  ✗ Ошибка: {e}")
        print("  PDF файлы будут созданы без русского текста")
        font_path = None

    all_files = []

    # Генерация документов
    print("1. Генерация документов:")
    print("-" * 50)

    for i, (title, category, content) in enumerate(DOCUMENTS, 1):
        base_name = f"doc_{i}_{category}"

        # .docx
        docx_path = os.path.join(DOCS_DIR, f"{base_name}.docx")
        generate_docx(docx_path, title, category, content)
        all_files.append(docx_path)
        print(f"  ✓ {base_name}.docx")

        # .doc
        doc_path = os.path.join(DOCS_DIR, f"{base_name}.doc")
        generate_doc(doc_path, title, category, content)
        all_files.append(doc_path)
        print(f"  ✓ {base_name}.doc")

        # .xlsx
        xlsx_path = os.path.join(DOCS_DIR, f"{base_name}.xlsx")
        generate_xlsx(xlsx_path, title, category, content)
        all_files.append(xlsx_path)
        print(f"  ✓ {base_name}.xlsx")

        # .xls
        xls_path = os.path.join(DOCS_DIR, f"{base_name}.xls")
        generate_xls(xls_path, title, category, content)
        all_files.append(xls_path)
        print(f"  ✓ {base_name}.xls")

        # .pdf (с русским шрифтом)
        pdf_path = os.path.join(DOCS_DIR, f"{base_name}.pdf")
        if font_path:
            try:
                generate_pdf(pdf_path, title, category, content, font_path)
                all_files.append(pdf_path)
                print(f"  ✓ {base_name}.pdf")
            except Exception as e:
                print(f"  ✗ {base_name}.pdf - ошибка: {e}")
        else:
            print(f"  ⚠ {base_name}.pdf - пропущен (нет шрифта)")

    print("-" * 50)
    print(f"  Всего сгенерировано документов: {len(all_files)}")

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
    print(f"   - Документы: {doc_count} файлов (.doc, .docx, .xls, .xlsx, .pdf)")
    print(f"   - Архивы: {arch_count} файлов (.zip, .rar, .7z)")
    print(
        f"   - Общий размер: {total_size / 1024:.1f} KB ({total_size / (1024*1024):.2f} MB)"
    )

    print("\n" + "=" * 70)
    print("✅ ГЕНЕРАЦИЯ ЗАВЕРШЕНА")
    print("=" * 70)
    print("\n🎯 ДАЛЕЕ ЗАПУСТИТЕ: python crawler.py")


if __name__ == "__main__":
    main()
