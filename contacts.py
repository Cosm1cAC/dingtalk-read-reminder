# -*- coding: utf-8 -*-
"""联系人手机号本地库（SQLite）：存一次，永久复用，不必每次提供。

数据文件: data/contacts.db
表: contacts(name 主键, mobile, updated_at)

excel_stats.load_people 会自动用这里的手机号补齐 Excel 里缺失的手机号。
"""
import os
import sqlite3
from datetime import datetime

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "contacts.db")


def _conn():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""CREATE TABLE IF NOT EXISTS contacts(
        name TEXT PRIMARY KEY,
        mobile TEXT NOT NULL,
        updated_at TEXT NOT NULL)""")
    return conn


def upsert(name: str, mobile: str):
    name = str(name).strip()
    mobile = str(mobile).strip().split(".")[0]
    if not name or not mobile:
        raise ValueError(f"姓名或手机号为空: {name!r} {mobile!r}")
    with _conn() as conn:
        conn.execute(
            "INSERT INTO contacts(name, mobile, updated_at) VALUES(?,?,?) "
            "ON CONFLICT(name) DO UPDATE SET mobile=excluded.mobile, updated_at=excluded.updated_at",
            (name, mobile, datetime.now().isoformat(timespec="seconds")))


def get(name: str) -> str:
    with _conn() as conn:
        row = conn.execute("SELECT mobile FROM contacts WHERE name=?", (str(name).strip(),)).fetchone()
    return row[0] if row else ""


def all_contacts() -> dict:
    with _conn() as conn:
        return {r[0]: r[1] for r in conn.execute("SELECT name, mobile FROM contacts")}


def count() -> int:
    with _conn() as conn:
        return conn.execute("SELECT COUNT(*) FROM contacts").fetchone()[0]


def import_pairs(pairs):
    """批量导入 [(姓名, 手机号), ...]，返回导入数量。"""
    for name, mobile in pairs:
        upsert(name, mobile)
    return len(pairs)


if __name__ == "__main__":
    import sys
    if len(sys.argv) == 3:
        upsert(sys.argv[1], sys.argv[2])
        print(f"[OK] {sys.argv[1]} -> {sys.argv[2]}")
    else:
        for name, mobile in sorted(all_contacts().items()):
            print(f"{name}\t{mobile}")
        print(f"共 {count()} 条")
