#!/usr/bin/env python3
"""
КРАУЛЕР ДЛЯ СКАНИРОВАНИЯ ФАЙЛОВ
- Сканирует хранилище
- Находит файлы всех форматов
- Парсит типы, названия и содержимое
- Сохраняет данные в CSV
Поддерживаемые форматы: .doc, .docx, .xls, .xlsx, .pdf, .zip, .rar, .7z
"""

import os
import csv
import zipfile
import subprocess
import tempfile
from pathlib import Path
import docx2txt
import pandas as pd
from PyPDF2 import PdfReader
import xlrd


class Crawler:
    def __init__(self, storage_dir="storage", output_csv="output/documents_index.csv"):
        self.storage_dir = Path(storage_dir)
        self.output_csv = output_csv
        self.results = []

        # Создаём output директорию
        os.makedirs(os.path.dirname(output_csv), exist_ok=True)

    def crawl(self):
        """Основной метод краулера - рекурсивное сканирование"""
        print("=" * 70)
        print("КРАУЛЕР")
        print("Сканирование хранилища и парсинг файлов")
        print("=" * 70)
        print(f"\n📁 Хранилище: {self.storage_dir}")
        print(f"📄 Выходной CSV: {self.output_csv}")
        print("\n" + "-" * 70)

        file_count = 0
        for file_path in self.storage_dir.rglob("*"):
            if file_path.is_file():
                file_count += 1
                self._process_file(file_path)

        print("-" * 70)
        print(f"\n✅ Сканирование завершено")
        print(f"   - Найдено файлов: {file_count}")
        print(f"   - Обработано: {len(self.results)}")

        return self.results

    def _process_file(self, file_path):
        """Обработка одного файла"""
        ext = file_path.suffix.lower()

        # Определяем тип файла и парсим содержимое
        if ext == ".docx":
            content = self._parse_docx(file_path)
            file_type = "docx"
        elif ext == ".doc":
            content = self._parse_doc(file_path)
            file_type = "doc"
        elif ext == ".xlsx":
            content = self._parse_xlsx(file_path)
            file_type = "xlsx"
        elif ext == ".xls":
            content = self._parse_xls(file_path)
            file_type = "xls"
        elif ext == ".pdf":
            content = self._parse_pdf(file_path)
            file_type = "pdf"
        elif ext in [".zip", ".rar", ".7z"]:
            content = self._parse_archive(file_path)
            file_type = ext[1:]  # zip, rar, 7z
        else:
            return  # Пропускаем неподдерживаемые форматы

        # Сохраняем результат
        self.results.append(
            {
                "file_path": str(file_path),
                "file_name": file_path.name,
                "file_type": file_type,
                "file_size": file_path.stat().st_size,
                "content": content,  # Полное содержимое
                "content_preview": content[:500] if len(content) > 500 else content,
            }
        )

        print(f"  ✓ {file_path.name} [{file_type}] - {len(content)} символов")

    def _parse_docx(self, file_path):
        """Парсинг .docx"""
        try:
            return docx2txt.process(file_path)
        except Exception as e:
            return f"[Ошибка парсинга DOCX: {e}]"

    def _parse_doc(self, file_path):
        """Парсинг .doc (как текстовый файл)"""
        try:
            with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                return f.read()
        except Exception as e:
            return f"[Ошибка парсинга DOC: {e}]"

    def _parse_xlsx(self, file_path):
        """Парсинг .xlsx"""
        try:
            df_dict = pd.read_excel(file_path, sheet_name=None)
            text_parts = []
            for sheet_name, df in df_dict.items():
                text_parts.append(f"=== Лист: {sheet_name} ===")
                text_parts.append(df.to_string())
            return "\n".join(text_parts)
        except Exception as e:
            return f"[Ошибка парсинга XLSX: {e}]"

    def _parse_xls(self, file_path):
        """Парсинг .xls (старый формат Excel)"""
        try:
            book = xlrd.open_workbook(file_path)
            text_parts = []
            for sheet_idx in range(book.nsheets):
                sheet = book.sheet_by_index(sheet_idx)
                text_parts.append(f"=== Лист {sheet_idx + 1}: {sheet.name} ===")
                for row in range(min(sheet.nrows, 100)):  # Ограничиваем 100 строками
                    row_values = [
                        str(sheet.cell_value(row, col)) for col in range(sheet.ncols)
                    ]
                    text_parts.append(" | ".join(row_values))
            return "\n".join(text_parts)
        except Exception as e:
            return f"[Ошибка парсинга XLS: {e}]"

    def _parse_pdf(self, file_path):
        """Парсинг PDF"""
        try:
            reader = PdfReader(file_path)
            text_parts = []
            for page in reader.pages:
                text_parts.append(page.extract_text())
            return "\n".join(text_parts)
        except Exception as e:
            return f"[Ошибка парсинга PDF: {e}]"

    def _parse_archive(self, file_path):
        """Парсинг архивов с рекурсивным извлечением вложенных файлов"""
        ext = file_path.suffix.lower()
        archive_content = [f"=== АРХИВ: {file_path.name} ==="]

        with tempfile.TemporaryDirectory() as tmpdir:
            try:
                # Распаковка
                if ext == ".zip":
                    with zipfile.ZipFile(file_path, "r") as zf:
                        zf.extractall(tmpdir)
                elif ext == ".rar":
                    subprocess.run(
                        f"unrar x -idq {file_path} {tmpdir}",
                        shell=True,
                        capture_output=True,
                    )
                elif ext == ".7z":
                    subprocess.run(
                        f"7z x -bsp0 -bso0 {file_path} -o{tmpdir}",
                        shell=True,
                        capture_output=True,
                    )

                # Парсим извлечённые файлы
                for extracted_file in Path(tmpdir).rglob("*"):
                    if extracted_file.is_file():
                        archive_content.append(
                            f"\n--- ВЛОЖЕНИЕ: {extracted_file.name} ---"
                        )
                        sub_ext = extracted_file.suffix.lower()

                        if sub_ext == ".docx":
                            archive_content.append(self._parse_docx(extracted_file))
                        elif sub_ext == ".doc":
                            archive_content.append(self._parse_doc(extracted_file))
                        elif sub_ext == ".xlsx":
                            archive_content.append(self._parse_xlsx(extracted_file))
                        elif sub_ext == ".xls":
                            archive_content.append(self._parse_xls(extracted_file))
                        elif sub_ext == ".pdf":
                            archive_content.append(self._parse_pdf(extracted_file))
                        else:
                            try:
                                with open(
                                    extracted_file,
                                    "r",
                                    encoding="utf-8",
                                    errors="ignore",
                                ) as f:
                                    archive_content.append(f.read()[:1000])
                            except:
                                archive_content.append(
                                    f"[Бинарный файл: {extracted_file.name}]"
                                )

            except Exception as e:
                archive_content.append(f"[Ошибка распаковки: {e}]")

        return "\n".join(archive_content)

    def save_to_csv(self):
        """Сохранение результатов в CSV файл (основное требование ТЗ)"""
        if not self.results:
            print("\n⚠ Нет данных для сохранения в CSV")
            return False

        with open(self.output_csv, "w", newline="", encoding="utf-8") as csvfile:
            fieldnames = [
                "file_path",
                "file_name",
                "file_type",
                "file_size",
                "content_preview",
            ]
            writer = csv.DictWriter(
                csvfile, fieldnames=fieldnames, delimiter=",", quotechar='"'
            )
            writer.writeheader()

            for result in self.results:
                writer.writerow(
                    {
                        "file_path": result["file_path"],
                        "file_name": result["file_name"],
                        "file_type": result["file_type"],
                        "file_size": result["file_size"],
                        "content_preview": result["content_preview"]
                        .replace("\n", " ")
                        .replace(",", ";"),
                    }
                )

        print(f"\n💾 CSV ФАЙЛ СОЗДАН: {self.output_csv}")
        print(f"   - Записей: {len(self.results)}")
        print(f"   - Размер: {os.path.getsize(self.output_csv) / 1024:.1f} KB")
        return True

    def save_full_texts(self):
        """Сохраняем полные тексты отдельно для импорта в БД"""
        import pickle

        with open("output/full_texts.pkl", "wb") as f:
            pickle.dump(self.results, f)
        print(f"💾 Полные тексты сохранены: output/full_texts.pkl")


def main():
    # Создаём и запускаем краулер
    crawler = Crawler()

    # Сканирование и парсинг
    results = crawler.crawl()

    # Сохранение в CSV (ОСНОВНОЕ ТРЕБОВАНИЕ)
    crawler.save_to_csv()

    # Сохраняем полные тексты для БД
    if results:
        crawler.save_full_texts()

    print("\n" + "=" * 70)
    print("✅ КРАУЛЕР ЗАВЕРШИЛ РАБОТУ")
    print(f"   CSV файл: {crawler.output_csv}")
    print("=" * 70)


if __name__ == "__main__":
    main()
