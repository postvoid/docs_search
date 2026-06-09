#!/usr/bin/env python3
"""
КРАУЛЕР ДЛЯ СКАНИРОВАНИЯ ФАЙЛОВ
Поддерживает: .doc, .docx, .xls, .xlsx, .pdf, .zip, .rar, .7z
"""

import os
import csv
import zipfile
import subprocess
import tempfile
from pathlib import Path
import docx2txt
import pandas as pd
import xlrd

try:
    from PyPDF2 import PdfReader

    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False


class Crawler:
    def __init__(self, storage_dir="storage", output_csv="output/documents_index.csv"):
        self.storage_dir = Path(storage_dir)
        self.output_csv = output_csv
        self.results = []
        self.stats = {
            "doc": 0,
            "docx": 0,
            "xls": 0,
            "xlsx": 0,
            "pdf": 0,
            "zip": 0,
            "rar": 0,
            "7z": 0,
            "other": 0,
        }
        os.makedirs(os.path.dirname(output_csv), exist_ok=True)

    def crawl(self):
        print("=" * 70)
        print("КРАУЛЕР")
        print("=" * 70)
        print(f"\n📁 Сканирование: {self.storage_dir}")
        print("\n" + "-" * 70)

        for file_path in self.storage_dir.rglob("*"):
            if file_path.is_file():
                self._process_file(file_path)

        print("-" * 70)
        self._print_stats()
        return self.results

    def _process_file(self, file_path):
        ext = file_path.suffix.lower()
        content = ""
        file_type = ext[1:]

        if ext == ".docx":
            content = self._parse_docx(file_path)
            self.stats["docx"] += 1
        elif ext == ".doc":
            content = self._parse_doc(file_path)
            self.stats["doc"] += 1
        elif ext == ".xlsx":
            content = self._parse_xlsx(file_path)
            self.stats["xlsx"] += 1
        elif ext == ".xls":
            content = self._parse_xls(file_path)
            self.stats["xls"] += 1
        elif ext == ".pdf" and PDF_AVAILABLE:
            content = self._parse_pdf(file_path)
            self.stats["pdf"] += 1
        elif ext == ".zip":
            content = self._parse_zip(file_path)
            self.stats["zip"] += 1
        elif ext == ".rar":
            content = self._parse_rar(file_path)
            self.stats["rar"] += 1
        elif ext == ".7z":
            content = self._parse_7z(file_path)
            self.stats["7z"] += 1
        else:
            self.stats["other"] += 1
            return

        if content:
            self.results.append(
                {
                    "file_path": str(file_path),
                    "file_name": file_path.name,
                    "file_type": file_type,
                    "file_size": file_path.stat().st_size,
                    "content": content,
                    "content_preview": content[:500],
                }
            )
            print(f"  ✓ {file_path.name} [{file_type}] - {len(content)} символов")

    def _parse_docx(self, path):
        try:
            return docx2txt.process(path) or "[Пустой DOCX]"
        except Exception as e:
            return f"[Ошибка: {e}]"

    def _parse_doc(self, path):
        try:
            with open(path, "r", encoding="utf-8", errors="ignore") as f:
                return f.read()
        except Exception as e:
            return f"[Ошибка: {e}]"

    def _parse_xlsx(self, path):
        try:
            df_dict = pd.read_excel(path, sheet_name=None)
            parts = []
            for sheet, df in df_dict.items():
                parts.append(f"=== Лист: {sheet} ===")
                parts.append(df.to_string())
            return "\n".join(parts)
        except Exception as e:
            return f"[Ошибка: {e}]"

    def _parse_xls(self, path):
        try:
            book = xlrd.open_workbook(path)
            parts = []
            for idx in range(book.nsheets):
                sheet = book.sheet_by_index(idx)
                parts.append(f"=== Лист {idx+1}: {sheet.name} ===")
                for row in range(min(sheet.nrows, 50)):
                    vals = [
                        str(sheet.cell_value(row, col)) for col in range(sheet.ncols)
                    ]
                    parts.append(" | ".join(vals))
            return "\n".join(parts)
        except Exception as e:
            return f"[Ошибка: {e}]"

    def _parse_pdf(self, path):
        try:
            reader = PdfReader(path)
            parts = []
            for i, page in enumerate(reader.pages, 1):
                text = page.extract_text()
                if text:
                    parts.append(f"=== Страница {i} ===")
                    parts.append(text)
            return "\n".join(parts) or "[PDF без текста]"
        except Exception as e:
            return f"[Ошибка: {e}]"

    def _parse_zip(self, path):
        content = [f"=== ZIP: {path.name} ==="]
        with tempfile.TemporaryDirectory() as tmp:
            try:
                with zipfile.ZipFile(path, "r") as zf:
                    zf.extractall(tmp)
                for f in Path(tmp).rglob("*"):
                    if f.is_file():
                        content.append(f"\n--- {f.name} ---")
                        content.append(self._parse_extracted(f))
            except Exception as e:
                content.append(f"[Ошибка: {e}]")
        return "\n".join(content)

    def _parse_rar(self, path):
        content = [f"=== RAR: {path.name} ==="]
        with tempfile.TemporaryDirectory() as tmp:
            try:
                subprocess.run(
                    f"unrar x -idq {path} {tmp}", shell=True, capture_output=True
                )
                for f in Path(tmp).rglob("*"):
                    if f.is_file():
                        content.append(f"\n--- {f.name} ---")
                        content.append(self._parse_extracted(f))
            except Exception as e:
                content.append(f"[Ошибка: {e}]")
        return "\n".join(content)

    def _parse_7z(self, path):
        content = [f"=== 7z: {path.name} ==="]
        with tempfile.TemporaryDirectory() as tmp:
            try:
                subprocess.run(
                    f"7z x -y -o{tmp} {path}", shell=True, capture_output=True
                )
                for f in Path(tmp).rglob("*"):
                    if f.is_file():
                        content.append(f"\n--- {f.name} ---")
                        content.append(self._parse_extracted(f))
            except Exception as e:
                content.append(f"[Ошибка: {e}]")
        return "\n".join(content)

    def _parse_extracted(self, path):
        ext = path.suffix.lower()
        if ext == ".docx":
            return self._parse_docx(path)
        elif ext == ".doc":
            return self._parse_doc(path)
        elif ext == ".xlsx":
            return self._parse_xlsx(path)
        elif ext == ".xls":
            return self._parse_xls(path)
        elif ext == ".pdf" and PDF_AVAILABLE:
            return self._parse_pdf(path)
        else:
            try:
                with open(path, "r", encoding="utf-8", errors="ignore") as f:
                    return f.read()[:1000]
            except:
                return f"[Файл: {path.name}]"

    def _print_stats(self):
        print(f"\n📊 СТАТИСТИКА:")
        for k, v in self.stats.items():
            if v > 0:
                print(f"   {k.upper()}: {v}")
        print(f"   Индексировано: {len(self.results)}")

    def save_to_csv(self):
        if not self.results:
            return False
        with open(self.output_csv, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(
                f,
                fieldnames=[
                    "file_path",
                    "file_name",
                    "file_type",
                    "file_size",
                    "content_preview",
                ],
            )
            writer.writeheader()
            for r in self.results:
                writer.writerow(
                    {
                        "file_path": r["file_path"],
                        "file_name": r["file_name"],
                        "file_type": r["file_type"],
                        "file_size": r["file_size"],
                        "content_preview": r["content_preview"].replace("\n", " ")[
                            :500
                        ],
                    }
                )
        print(f"\n💾 CSV: {self.output_csv} ({len(self.results)} записей)")
        return True

    def save_full_texts(self):
        import pickle

        with open("output/full_texts.pkl", "wb") as f:
            pickle.dump(self.results, f)
        print(f"💾 Полные тексты: output/full_texts.pkl")


def main():
    crawler = Crawler()
    crawler.crawl()
    crawler.save_to_csv()
    crawler.save_full_texts()
    print("\n✅ КРАУЛЕР ЗАВЕРШЁН")


if __name__ == "__main__":
    main()
