"""
Менеджер базы данных SQLite для трекинга скачанных файлов.

Функционал:
  - Сохранение информации о каждом скачанном файле
  - Предварительная проверка дубликатов (по URL и MD5) перед скачиванием
  - Учёт суммарного размера скачанных файлов
  - Статистика по базам, блокам, ключевым словам
"""

import hashlib
import logging
import os
import sqlite3
from datetime import datetime
from pathlib import Path

logger = logging.getLogger("parser")


class DatabaseManager:
    """Менеджер SQLite-базы для учёта скачанных файлов."""

    DB_FILE = "downloaded_files.db"

    def __init__(self, db_dir=None):
        if db_dir is None:
            db_dir = Path(".")
        db_dir = Path(db_dir)
        db_dir.mkdir(parents=True, exist_ok=True)
        self.db_path = db_dir / self.DB_FILE
        self._conn = None
        self._init_db()

    # ──────────────────────────────────────────────────────────
    #  Инициализация
    # ──────────────────────────────────────────────────────────

    def _init_db(self):
        """Создаёт таблицы, если их ещё нет."""
        self._conn = sqlite3.connect(str(self.db_path))
        self._conn.row_factory = sqlite3.Row

        self._conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS downloaded_files (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                url             TEXT NOT NULL UNIQUE,
                md5_hash        TEXT UNIQUE,
                title           TEXT,
                authors         TEXT,
                year            INTEGER,
                journal         TEXT,
                doi             TEXT,
                keywords        TEXT,
                abstract        TEXT,
                file_path       TEXT,
                file_size       INTEGER DEFAULT 0,
                file_name       TEXT,
                block_number    INTEGER,
                query_keyword   TEXT,
                download_date   TEXT,
                status          TEXT DEFAULT 'downloaded'
            );

            CREATE INDEX IF NOT EXISTS idx_url   ON downloaded_files(url);
            CREATE INDEX IF NOT EXISTS idx_md5   ON downloaded_files(md5_hash);
            CREATE INDEX IF NOT EXISTS idx_block ON downloaded_files(block_number);
            CREATE INDEX IF NOT EXISTS idx_query ON downloaded_files(query_keyword);

            CREATE TABLE IF NOT EXISTS download_stats (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                stat_key        TEXT UNIQUE,
                stat_value      TEXT
            );
            """
        )
        self._conn.commit()
        logger.info("БД инициализирована: %s", self.db_path)

    # ──────────────────────────────────────────────────────────
    #  Предварительная проверка дубликатов (до скачивания)
    # ──────────────────────────────────────────────────────────

    def is_url_downloaded(self, url):
        """Проверяет, был ли файл с данным URL уже скачан."""
        row = self._conn.execute(
            "SELECT 1 FROM downloaded_files WHERE url = ? AND status = 'downloaded'",
            (url,),
        ).fetchone()
        return row is not None

    def is_md5_known(self, md5_hash):
        """Проверяет, есть ли уже файл с таким MD5."""
        if not md5_hash:
            return False
        row = self._conn.execute(
            "SELECT 1 FROM downloaded_files WHERE md5_hash = ? AND status = 'downloaded'",
            (md5_hash,),
        ).fetchone()
        return row is not None

    def check_duplicate_before_download(self, url):
        """
        Предварительная проверка перед скачиванием.
        Возвращает (is_duplicate, existing_record).
        """
        if self.is_url_downloaded(url):
            row = self._conn.execute(
                "SELECT * FROM downloaded_files WHERE url = ?",
                (url,),
            ).fetchone()
            logger.info("Дубликат по URL: %s", url)
            return True, dict(row) if row else None
        return False, None

    # ──────────────────────────────────────────────────────────
    #  Сохранение записи о скачанном файле
    # ──────────────────────────────────────────────────────────

    @staticmethod
    def compute_md5(file_path):
        """Вычисляет MD5-хеш файла."""
        md5 = hashlib.md5()
        with open(file_path, "rb") as f:
            for chunk in iter(lambda: f.read(8192), b""):
                md5.update(chunk)
        return md5.hexdigest()

    def save_file_record(
        self,
        url,
        file_path,
        title="",
        authors="",
        year=None,
        journal="",
        doi="",
        keywords="",
        abstract="",
        block_number=None,
        query_keyword="",
        status="downloaded",
    ):
        """
        Сохраняет запись о скачанном файле в БД.
        Вычисляет MD5 и размер автоматически.
        Возвращает (md5_hash, file_size) или (None, 0) при ошибке.
        """
        file_path = Path(file_path)
        md5_hash = None
        file_size = 0

        if file_path.exists():
            md5_hash = self.compute_md5(file_path)
            file_size = file_path.stat().st_size

            # Проверяем дубликат по MD5 (файл с другим URL, но тем же содержимым)
            if self.is_md5_known(md5_hash):
                existing = self._conn.execute(
                    "SELECT * FROM downloaded_files WHERE md5_hash = ? AND status = 'downloaded'",
                    (md5_hash,),
                ).fetchone()
                logger.warning(
                    "Дубликат по MD5: %s (уже скачан как: %s)",
                    file_path.name,
                    existing["file_name"] if existing else "unknown",
                )
                self._conn.execute(
                    """INSERT OR REPLACE INTO downloaded_files
                       (url, md5_hash, title, authors, year, journal, doi,
                        keywords, abstract, file_path, file_size, file_name,
                        block_number, query_keyword, download_date, status)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                    (
                        url, md5_hash, title, authors, year, journal, doi,
                        keywords, abstract, str(file_path), file_size,
                        file_path.name, block_number, query_keyword,
                        datetime.now().isoformat(), "duplicate",
                    ),
                )
                self._conn.commit()
                return md5_hash, file_size
        else:
            status = "missing"

        self._conn.execute(
            """INSERT OR REPLACE INTO downloaded_files
               (url, md5_hash, title, authors, year, journal, doi,
                keywords, abstract, file_path, file_size, file_name,
                block_number, query_keyword, download_date, status)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                url, md5_hash, title, authors, year, journal, doi,
                keywords, abstract, str(file_path), file_size,
                file_path.name, block_number, query_keyword,
                datetime.now().isoformat(), status,
            ),
        )
        self._conn.commit()
        logger.info(
            "Запись сохранена в БД: %s (MD5: %s, %d байт)",
            file_path.name, md5_hash, file_size,
        )
        return md5_hash, file_size

    # ──────────────────────────────────────────────────────────
    #  Учёт суммарного размера
    # ──────────────────────────────────────────────────────────

    def get_total_size(self):
        """Возвращает суммарный размер всех скачанных файлов в байтах."""
        row = self._conn.execute(
            "SELECT COALESCE(SUM(file_size), 0) AS total "
            "FROM downloaded_files WHERE status = 'downloaded'",
        ).fetchone()
        return row["total"] if row else 0

    def get_total_size_mb(self):
        """Возвращает суммарный размер в мегабайтах."""
        return self.get_total_size() / (1024 * 1024)

    def get_total_size_gb(self):
        """Возвращает суммарный размер в гигабайтах."""
        return self.get_total_size() / (1024 * 1024 * 1024)

    def get_downloaded_count(self):
        """Возвращает количество уникальных скачанных файлов."""
        row = self._conn.execute(
            "SELECT COUNT(*) AS cnt FROM downloaded_files WHERE status = 'downloaded'",
        ).fetchone()
        return row["cnt"] if row else 0

    def get_duplicate_count(self):
        """Возвращает количество отбракованных дубликатов."""
        row = self._conn.execute(
            "SELECT COUNT(*) AS cnt FROM downloaded_files WHERE status = 'duplicate'",
        ).fetchone()
        return row["cnt"] if row else 0

    # ──────────────────────────────────────────────────────────
    #  Статистика
    # ──────────────────────────────────────────────────────────

    def get_stats(self):
        """Возвращает сводную статистику."""
        total_size = self.get_total_size()
        return {
            "total_files": self.get_downloaded_count(),
            "duplicates": self.get_duplicate_count(),
            "total_size_bytes": total_size,
            "total_size_mb": total_size / (1024 * 1024),
            "total_size_gb": total_size / (1024 * 1024 * 1024),
        }

    def get_stats_by_keyword(self):
        """Статистика по ключевым словам."""
        rows = self._conn.execute(
            """SELECT query_keyword,
                      COUNT(*) AS cnt,
                      SUM(file_size) AS total_size
               FROM downloaded_files
               WHERE status = 'downloaded'
               GROUP BY query_keyword
               ORDER BY cnt DESC""",
        ).fetchall()
        return [
            {
                "keyword": r["query_keyword"],
                "count": r["cnt"],
                "size_mb": (r["total_size"] or 0) / (1024 * 1024),
            }
            for r in rows
        ]

    def get_stats_by_block(self):
        """Статистика по блокам директорий."""
        rows = self._conn.execute(
            """SELECT block_number,
                      COUNT(*) AS cnt,
                      SUM(file_size) AS total_size
               FROM downloaded_files
               WHERE status = 'downloaded'
               GROUP BY block_number
               ORDER BY block_number""",
        ).fetchall()
        return [
            {
                "block": r["block_number"],
                "count": r["cnt"],
                "size_mb": (r["total_size"] or 0) / (1024 * 1024),
            }
            for r in rows
        ]

    def get_files_in_block(self, block_number):
        """Возвращает все файлы в указанном блоке."""
        rows = self._conn.execute(
            "SELECT * FROM downloaded_files WHERE block_number = ? ORDER BY id",
            (block_number,),
        ).fetchall()
        return [dict(r) for r in rows]

    def get_known_urls(self):
        """Возвращает множество всех URL, которые уже были обработаны."""
        rows = self._conn.execute(
            "SELECT url FROM downloaded_files",
        ).fetchall()
        return {r["url"] for r in rows}

    def get_known_md5_hashes(self):
        """Возвращает множество всех известных MD5-хешей."""
        rows = self._conn.execute(
            "SELECT md5_hash FROM downloaded_files WHERE md5_hash IS NOT NULL",
        ).fetchall()
        return {r["md5_hash"] for r in rows}

    def close(self):
        """Закрывает соединение с БД."""
        if self._conn:
            self._conn.close()
            self._conn = None
