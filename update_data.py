# -*- coding: utf-8 -*-
"""每日数据更新：从部长助手导出的员工明细中筛出指定日期（「日」类型）数据，
写入 data/当前统计.xlsx（含数据日期 meta），并逐人交叉验证。

用法:
    python update_data.py <明细.xlsx> [YYYY-MM-DD]
不传日期时默认取表内最大的「日」类型日期。
"""
import datetime
import os
import sys
import warnings
from collections import Counter

warnings.filterwarnings("ignore")

from openpyxl import Workbook, load_workbook

from excel_stats import COLUMNS, load_data_date, save_data_date

TARGET = os.path.join("data", "当前统计.xlsx")


def extract(src_path: str, date_str: str = None) -> tuple:
    wb = load_workbook(src_path, data_only=True)
    ws = wb.active
    rows = [r for r in ws.iter_rows(min_row=2, values_only=True) if r and r[0]]
    dist = Counter((str(r[0])[:10], str(r[1]).strip()) for r in rows)
    print("日期/类型分布:", dict(dist))

    day_rows = [r for r in rows if str(r[1]).strip() == "日"]
    if not date_str:
        date_str = max(str(r[0])[:10] for r in day_rows)
    picked = [r for r in day_rows if str(r[0])[:10] == date_str]
    if not picked:
        raise SystemExit(f"{src_path} 里没有 {date_str} 的「日」类型数据")

    data = {}
    for r in picked:
        data[str(r[8]).strip()] = (str(r[9] or "").strip(), int(r[10] or 0), int(r[11] or 0))
    return date_str, data


def write_target(date_str: str, data: dict, target: str = TARGET) -> str:
    tmp = target.replace(".xlsx", "_new.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "已读统计"
    ws.append(COLUMNS)
    for name, (role, total, read) in data.items():
        ws.append([name, "", role, total, read])
    wb.save(tmp)
    save_data_date(tmp, datetime.date.fromisoformat(date_str))
    try:
        os.replace(tmp, target)
        return target
    except PermissionError:
        print(f"[WARN] {target} 被占用，新数据暂存在 {tmp}")
        return tmp


def verify(path: str, data: dict) -> None:
    ws = load_workbook(path, data_only=True)["已读统计"]
    dst = {str(r[0]).strip(): (str(r[2] or "").strip(), int(r[3]), int(r[4]))
           for r in ws.iter_rows(min_row=2, values_only=True) if r and r[0]}
    bad = [(n, v, dst.get(n)) for n, v in data.items() if dst.get(n) != v]
    print(f"交叉验证: {len(data)} 人 vs {len(dst)} 人 | 差异 {len(bad)} | meta 日期 {load_data_date(path)}")
    for b in bad[:10]:
        print("  不一致:", b)


def main():
    if len(sys.argv) < 2:
        raise SystemExit(__doc__)
    src = sys.argv[1]
    date_str = sys.argv[2] if len(sys.argv) > 2 else None
    date_str, data = extract(src, date_str)
    print(f"{date_str}（日类型）人数: {len(data)} | 总消息量 {sum(v[1] for v in data.values())} "
          f"| 总已读 {sum(v[2] for v in data.values())}")
    path = write_target(date_str, data)
    verify(path, data)


if __name__ == "__main__":
    main()
