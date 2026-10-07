"""
Менеджер спецификаций документов.

Обновления:
  - Спецификации хранятся в той же директории, что и PDF-файлы (в блочных директориях)
  - Интеграция с DatabaseManager для проверки дубликатов
  - Блочная структура: block_001/, block_002/, ... по 10 документов
"""

import hashlib
import json
import logging
from pathlib import Path

logger = logging.getLogger("parser")


class SpecManager:
    """
    Управляет спецификациями документов и блочными директориями.

    Структура:
        data/
          block_001/
            2024 - Название статьи - a1b2c3d4.pdf
            2024 - Название статьи - a1b2c3d4.json    ← спецификация рядом с PDF
          block_002/
            ...
    """

    DOCS_PER_BLOCK = 10
    BASE_DATA_DIR = Path("data")

    def __init__(self, db_manager=None):
        self.base_dir = self.BASE_DATA_DIR
        self.db_manager = db_manager
        self.docs_per_block = self.DOCS_PER_BLOCK
        self.known_hashes = set()
        self._current_block = None
        self._block_doc_count = 0

        # Если есть БД — загружаем хеши из неё
        if self.db_manager is not None:
            self.known_hashes = self.db_manager.get_known_md5_hashes()

    def init_dirs(self):
        """Создаёт базовую директорию и первый блок."""
        self.base_dir.mkdir(parents=True, exist_ok=True)

        # Если БД есть — синхронизируем known_hashes
        if self.db_manager is not None:
            self.known_hashes = self.db_manager.get_known_md5_hashes()

        # Определяем текущий блок
        self._current_block = self.get_next_block_number()
        self.ensure_block_dir(self._current_block)
        self._update_block_doc_count()

        logger.info(
            "SpecManager: базовая директория %s, текущий блок %d, документов в блоке %d",
            self.base_dir, self._current_block, self._block_doc_count,
        )

    # ──────────────────────────────────────────────────────────
    #  Блочные директории
    # ──────────────────────────────────────────────────────────

    def get_block_dir(self, block_number):
        """Возвращает путь к директории блока."""
        return self.base_dir / f"block_{block_number:03d}"

    def ensure_block_dir(self, block_number):
        """Создаёт директорию блока, если её нет."""
        block_dir = self.get_block_dir(block_number)
        block_dir.mkdir(parents=True, exist_ok=True)
        return block_dir

    def get_next_block_number(self):
        """
        Определяет следующий номер блока.
        Приоритет: БД → файловая система → 1.
        """
        if self.db_manager is not None:
            row = None
            try:
                row = self.db_manager._conn.execute(
                    "SELECT MAX(block_number) AS max_block "
                    "FROM downloaded_files WHERE status = 'downloaded'",
                ).fetchone()
            except Exception:
                pass
            if row and row["max_block"] is not None:
                max_block = row["max_block"]
                # Считаем количество документов в последнем блоке
                count_row = self.db_manager._conn.execute(
                    "SELECT COUNT(*) AS cnt FROM downloaded_files "
                    "WHERE block_number = ? AND status = 'downloaded'",
                    (max_block,),
                ).fetchone()
                if count_row and count_row["cnt"] >= self.docs_per_block:
                    return max_block + 1
                return max_block

        # По файловой системе
        if not self.base_dir.exists():
            return 1

        blocks = sorted(self.base_dir.glob("block_*"))
        if not blocks:
            return 1

        last_block = blocks[-1]
        block_num = int(last_block.name.split("_")[1])
        docs_in_block = len(list(last_block.glob("*.pdf")))
        if docs_in_block >= self.docs_per_block:
            return block_num + 1
        return block_num

    def _update_block_doc_count(self):
        """Обновляет счётчик документов в текущем блоке."""
        if self._current_block is None:
            self._block_doc_count = 0
            return

        block_dir = self.get_block_dir(self._current_block)
        if block_dir.exists():
            self._block_doc_count = len(list(block_dir.glob("*.pdf")))
        else:
            self._block_doc_count = 0

    def maybe_rotate_block(self):
        """
        Проверяет, заполнен ли текущий блок (10 документов).
        Если заполнен — переключается на следующий.
        Возвращает True, если произошла ротация.
        """
        self._update_block_doc_count()
        if self._block_doc_count >= self.docs_per_block:
            self._current_block += 1
            self.ensure_block_dir(self._current_block)
            self._block_doc_count = 0
            logger.info("Ротация: новый блок block_%03d", self._current_block)
            return True
        return False

    # ──────────────────────────────────────────────────────────
    #  Контроль дубликатов
    # ──────────────────────────────────────────────────────────

    @staticmethod
    def compute_md5(file_path):
        """Вычисляет MD5-хеш файла."""
        md5 = hashlib.md5()
        with open(file_path, "rb") as f:
            for chunk in iter(lambda: f.read(8192), b""):
                md5.update(chunk)
        return md5.hexdigest()

    def is_duplicate(self, file_path):
        """
        Проверяет, является ли файл дубликатом по MD5.
        Использует БД, если доступна; иначе — known_hashes.
        """
        file_path = Path(file_path)
        if not file_path.exists():
            return False

        md5 = self.compute_md5(file_path)

        if md5 in self.known_hashes:
            return True

        if self.db_manager is not None:
            if self.db_manager.is_md5_known(md5):
                self.known_hashes.add(md5)
                return True

        return False

    # ──────────────────────────────────────────────────────────
    #  Создание спецификаций (в той же директории, что и PDF)
    # ──────────────────────────────────────────────────────────

    def create_spec(
        self,
        file_path,
        title="",
        authors="",
        abstract="",
        url="",
        year=None,
        doi="",
        keywords="",
        block_number=None,
    ):
        """
        Создаёт JSON-спецификацию рядом с PDF-файлом.
        Имя спецификации = имя PDF без расширения + .json
        """
        file_path = Path(file_path)
        if not file_path.exists():
            logger.error("Файл не найден: %s", file_path)
            return None

        if block_number is None:
            block_number = self._current_block

        block_dir = self.get_block_dir(block_number)
        block_dir.mkdir(parents=True, exist_ok=True)

        # Спецификация рядом с PDF, в той же директории
        spec_name = file_path.stem + ".json"
        spec_path = block_dir / spec_name

        md5 = self.compute_md5(file_path)
        file_size = file_path.stat().st_size

        spec_data = {
            "file_name": file_path.name,
            "file_path": str(file_path),
            "file_size": file_size,
            "md5_hash": md5,
            "title": title,
            "authors": authors,
            "year": year,
            "doi": doi,
            "journal": "",
            "keywords": keywords,
            "abstract": abstract,
            "url": url,
            "block_number": block_number,
            "spec_version": "2.0",
        }

        try:
            spec_path.write_text(
                json.dumps(spec_data, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
            logger.info("Спецификация создана: %s", spec_path)

            # Добавляем MD5 в known_hashes
            self.known_hashes.add(md5)

            return spec_path

        except Exception as e:
            logger.error("Ошибка создания спецификации: %s", e)
            return None

    # ──────────────────────────────────────────────────────────
    #  Статистика
    # ──────────────────────────────────────────────────────────

    def get_stats(self):
        """Возвращает сводную статистику."""
        if self.db_manager is not None:
            db_stats = self.db_manager.get_stats()
        else:
            db_stats = {"total_files": 0, "total_size_mb": 0}

        # По файловой системе
        if not self.base_dir.exists():
            return {
                "total_docs": 0,
                "total_size_mb": 0,
                "blocks": 0,
                "block_numbers": [],
            }

        total_size = 0
        total_docs = 0
        block_numbers = []

        for block_dir in sorted(self.base_dir.glob("block_*")):
            try:
                block_num = int(block_dir.name.split("_")[1])
            except (ValueError, IndexError):
                continue
            block_numbers.append(block_num)

            for pdf in block_dir.glob("*.pdf"):
                total_docs += 1
                try:
                    total_size += pdf.stat().st_size
                except OSError:
                    pass

        return {
            "total_docs": total_docs,
            "total_size_mb": total_size / (1024 * 1024),
            "blocks": len(block_numbers),
            "block_numbers": block_numbers,
            "db_files": db_stats["total_files"],
            "db_size_mb": db_stats.get("total_size_mb", 0),
        }

    def get_block_dir_for_new_file(self):
        """Возвращает директорию для нового файла (с учётом ротации)."""
        self.maybe_rotate_block()
        return self.get_block_dir(self._current_block)

    @property
    def current_block(self):
        return self._current_block
