from __future__ import annotations

import os
import sqlite3
import sys
from pathlib import Path

import psycopg
from psycopg import sql


SOURCE_DEFAULT = Path(__file__).resolve().parent.parent / ".sistem" / "data" / "imei_pos.db"
TABLES = ("purchase_batches", "barang", "penjualan", "import_log", "audit_log")


def target_url() -> str:
    value = os.environ.get("DATABASE_URL_UNPOOLED") or os.environ.get("DATABASE_URL")
    if not value:
        raise RuntimeError("DATABASE_URL_UNPOOLED atau DATABASE_URL belum tersedia.")
    return value


def sqlite_rows(source: sqlite3.Connection, table_name: str) -> tuple[list[str], list[tuple[object, ...]]]:
    columns = [row["name"] for row in source.execute(f"PRAGMA table_info({table_name})")]
    rows = [tuple(row[column] for column in columns) for row in source.execute(f"SELECT * FROM {table_name}")]
    return columns, rows


def upsert_rows(target: psycopg.Connection, table_name: str, columns: list[str], rows: list[tuple[object, ...]]) -> None:
    if not rows:
        return
    primary_key = "imei" if table_name == "barang" else "id"
    update_columns = [column for column in columns if column != primary_key]
    statement = sql.SQL("INSERT INTO {table} ({columns}) VALUES ({values}) ON CONFLICT ({pk}) DO UPDATE SET {updates}").format(
        table=sql.Identifier(table_name),
        columns=sql.SQL(", ").join(map(sql.Identifier, columns)),
        values=sql.SQL(", ").join(sql.Placeholder() for _ in columns),
        pk=sql.Identifier(primary_key),
        updates=sql.SQL(", ").join(
            sql.SQL("{column}=EXCLUDED.{column}").format(column=sql.Identifier(column)) for column in update_columns
        ),
    )
    with target.cursor() as cursor:
        cursor.executemany(statement, rows)


def sync_sequence(target: psycopg.Connection, table_name: str) -> None:
    if table_name == "barang":
        return
    target.execute(
        sql.SQL(
            "SELECT setval(pg_get_serial_sequence({table_literal}, 'id'), "
            "GREATEST(COALESCE((SELECT MAX(id) FROM {table}), 1), 1), "
            "EXISTS(SELECT 1 FROM {table}))"
        ).format(table_literal=sql.Literal(table_name), table=sql.Identifier(table_name))
    )


def main() -> None:
    source_path = Path(sys.argv[1]).resolve() if len(sys.argv) > 1 else SOURCE_DEFAULT.resolve()
    if not source_path.exists():
        raise FileNotFoundError(f"Database SQLite tidak ditemukan: {source_path}")

    url = target_url()
    os.environ["DATABASE_URL"] = url
    from app import init_db

    init_db()
    source = sqlite3.connect(source_path)
    source.row_factory = sqlite3.Row

    try:
        with psycopg.connect(url) as target:
            for table_name in TABLES:
                columns, rows = sqlite_rows(source, table_name)
                upsert_rows(target, table_name, columns, rows)
                sync_sequence(target, table_name)

            for table_name in TABLES:
                source_count = source.execute(f"SELECT COUNT(*) FROM {table_name}").fetchone()[0]
                target_count = target.execute(
                    sql.SQL("SELECT COUNT(*) FROM {}").format(sql.Identifier(table_name))
                ).fetchone()[0]
                if target_count != source_count:
                    raise RuntimeError(
                        f"Validasi gagal untuk {table_name}: SQLite={source_count}, Neon={target_count}"
                    )
                print(f"{table_name}: {target_count} baris")
    finally:
        source.close()

    print("Migrasi data selesai dan jumlah seluruh tabel cocok.")


if __name__ == "__main__":
    main()
