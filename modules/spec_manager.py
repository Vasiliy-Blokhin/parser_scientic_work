"""
Менеджер спецификаций документов.

Структура (категории → блоки по 10 документов):
    data/
      Научные работы/
        block_001/
          2025 - Название - a1b2c3d4.pdf
          2025 - Название - a1b2c3d4.json    ← спецификация рядом с PDF
        block_002/
      Войсковой ремонт/
        block_001/
      ...

Нумерация блоков — независимая для каждой категории.
"""

import hashlib
import json
import logging
import re
from pathlib import Path

from modules.settings import DOCS_PER_BLOCK

logger = logging.getLogger("parser")


class SpecManager:
    """Управляет спецификациями, категориями и блочными директориями."""

    BASE_DATA_DIR = Path("data")

    def __init__(self, db_manager=None):
        self.base_dir = self.BASE_DATA_DIR
        self.db_manager = db_manager
        self.docs_per_block = DOCS_PER_BLOCK
        self.known_hashes = set()
        # {категория: номер текущего блока}
        self._current_blocks = {}

        if self.db_manager is not None:
            self.known_hashes = self.db_manager.get_known_md5_hashes()

    def init_dirs(self):
        """Создаёт базовую директорию."""
        self.base_dir.mkdir(parents=True, exist_ok=True)
        if self.db_manager is not None:
            self.known_hashes = self.db_manager.get_known_md5_hashes()
        logger.info("SpecManager: базовая директория %s", self.base_dir)

    # ──────────────────────────────────────────────────────────
    #  Категории и блочные директории
    # ──────────────────────────────────────────────────────────

    @staticmethod
    def safe_category_name(category):
        """Имя категории, безопасное для файловой системы."""
        name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", category or "")
        name = re.sub(r"\s+", " ", name).strip(" .")
        return name or "Без категории"

    def get_category_dir(self, category):
        return self.base_dir / self.safe_category_name(category)

    def get_block_dir(self, category, block_number):
        return self.get_category_dir(category) / f"block_{block_number:03d}"

    def ensure_block_dir(self, category, block_number):
        block_dir = self.get_block_dir(category, block_number)
        block_dir.mkdir(parents=True, exist_ok=True)
        return block_dir

    def _count_docs(self, category, block_number):
        block_dir = self.get_block_dir(category, block_number)
        if not block_dir.exists():
            return 0
        return len(list(block_dir.glob("*.pdf")))

    def _detect_block(self, category):
        """Определяет текущий блок категории по файловой системе."""
        cat_dir = self.get_category_dir(category)
        if not cat_dir.exists():
            return 1
        numbers = []
        for d in cat_dir.glob("block_*"):
            try:
                numbers.append(int(d.name.split("_")[1]))
            except (ValueError, IndexError):
                continue
        if not numbers:
            return 1
        last = max(numbers)
        if self._count_docs(category, last) >= self.docs_per_block:
            return last + 1
        return last

    def get_current_block(self, category):
        """Текущий блок категории (при первом обращении — определяется)."""
        if category not in self._current_blocks:
            self._current_blocks[category] = self._detect_block(category)
            self.ensure_block_dir(category, self._current_blocks[category])
        return self._current_blocks[category]

    def maybe_rotate_block(self, category):
        """
        Если блок категории заполнен (10 документов) — открывает следующий.
        Возвращает True, если произошла ротация.
        """
        block = self.get_current_block(category)
        if self._count_docs(category, block) >= self.docs_per_block:
            block += 1
            self._current_blocks[category] = block
            self.ensure_block_dir(category, block)
            logger.info("Ротация [%s]: новый блок block_%03d", category, block)
            return True
        return False

    def get_block_dir_for_new_file(self, category):
        """Директория для нового файла (с учётом ротации)."""
        self.maybe_rotate_block(category)
        return self.get_block_dir(category, self.get_current_block(category))

    # ──────────────────────────────────────────────────────────
    #  Контроль дубликатов
    # ──────────────────────────────────────────────────────────

    @staticmethod
    def compute_md5(file_path):
        md5 = hashlib.md5()
        with open(file_path, "rb") as f:
            for chunk in iter(lambda: f.read(8192), b""):
                md5.update(chunk)
        return md5.hexdigest()

    def is_duplicate(self, file_path):
        """Дубликат по MD5 (known_hashes + БД)."""
        file_path = Path(file_path)
        if not file_path.exists():
            return False

        md5 = self.compute_md5(file_path)
        if md5 in self.known_hashes:
            return True

        if self.db_manager is not None and self.db_manager.is_md5_known(md5):
            self.known_hashes.add(md5)
            return True
        return False

    # ──────────────────────────────────────────────────────────
    #  Создание спецификаций (рядом с PDF)
    # ──────────────────────────────────────────────────────────

    def create_spec(
        self,
        file_path,
        category="",
        title="",
        authors="",
        abstract="",
        url="",
        year=None,
        doi="",
        journal="",
        keywords="",
        block_number=None,
    ):
        """JSON-спецификация рядом с PDF (имя = имя PDF + .json)."""
        file_path = Path(file_path)
        if not file_path.exists():
            logger.error("Файл не найден: %s", file_path)
            return None

        if block_number is None:
            block_number = self.get_current_block(category)

        spec_path = file_path.with_suffix(".json")
        md5 = self.compute_md5(file_path)

        spec_data = {
            "file_name": file_path.name,
            "file_path": str(file_path),
            "file_size": file_path.stat().st_size,
            "md5_hash": md5,
            "category": category,
            "title": title,
            "authors": authors,
            "year": year,
            "doi": doi,
            "journal": journal,
            "keywords": keywords,
            "abstract": abstract,
            "url": url,
            "block_number": block_number,
            "spec_version": "3.0",
        }

        try:
            spec_path.write_text(
                json.dumps(spec_data, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
            logger.info("Спецификация создана: %s", spec_path)
            self.known_hashes.add(md5)
            return spec_path
        except Exception as e:
            logger.error("Ошибка создания спецификации: %s", e)
            return None

    # ──────────────────────────────────────────────────────────
    #  Статистика
    # ──────────────────────────────────────────────────────────

    def get_stats(self):
        """Статистика по файловой системе: всего и по категориям."""
        empty = {"total_docs": 0, "total_size_mb": 0, "categories": {}}
        if not self.base_dir.exists():
            return empty

        total_size = 0
        total_docs = 0
        categories = {}

        for cat_dir in sorted(p for p in self.base_dir.iterdir() if p.is_dir()):
            docs = 0
            size = 0
            blocks = 0
            for block_dir in sorted(cat_dir.glob("block_*")):
                blocks += 1
                for pdf in block_dir.glob("*.pdf"):
                    docs += 1
                    try:
                        size += pdf.stat().st_size
                    except OSError:
                        pass
            categories[cat_dir.name] = {
                "docs": docs,
                "blocks": blocks,
                "size_mb": size / (1024 * 1024),
            }
            total_docs += docs
            total_size += size

        return {
            "total_docs": total_docs,
            "total_size_mb": total_size / (1024 * 1024),
            "categories": categories,
        }
