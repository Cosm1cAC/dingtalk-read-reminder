# -*- coding: utf-8 -*-
"""读取已读统计 Excel 并计算统计数据。

Excel 模板列（见 main.py make-template 生成）：
    姓名 | 手机号 | 职位 | 应读条数 | 已读条数

分类规则（与示例图一致）：
    - 使用率 = 已读条数 / 应读条数
    - 未使用：使用率 < 50%
    - 已全部读完：已读条数 == 应读条数
    - 部分已读：使用率 >= 50% 且未读完（按完成率降序）
    - 整体使用率 = sum(已读) / sum(应读)
"""
import os
from dataclasses import dataclass, field

from openpyxl import Workbook, load_workbook

COLUMNS = ["姓名", "手机号", "职位", "应读条数", "已读条数"]
THRESHOLD = 0.5  # 未使用阈值


@dataclass
class Person:
    name: str
    mobile: str
    role: str = ""
    total: int = 0     # 应读条数
    read: int = 0      # 已读条数

    @property
    def rate(self) -> float:
        return self.read / self.total if self.total else 0.0

    @property
    def rate_str(self) -> str:
        return f"{self.rate * 100:.2f}%"


@dataclass
class Stats:
    all_people: list = field(default_factory=list)

    @property
    def overall_rate(self) -> float:
        total = sum(p.total for p in self.all_people)
        read = sum(p.read for p in self.all_people)
        return read / total if total else 0.0

    @property
    def unused(self) -> list:
        return [p for p in self.all_people if p.rate < THRESHOLD]

    @property
    def finished(self) -> list:
        return [p for p in self.all_people if p.total > 0 and p.read >= p.total]

    @property
    def partial(self) -> list:
        return [p for p in self.all_people
                if THRESHOLD <= p.rate < 1.0 and p.total > 0]

    @property
    def partial_sorted(self) -> list:
        return sorted(self.partial, key=lambda p: p.rate, reverse=True)


def load_people(excel_path: str) -> list:
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"找不到数据文件: {excel_path}，请先执行 main.py make-template 生成模板并填写数据")
    # 联系人库：Excel 里手机号缺失时自动从这里补
    try:
        import contacts
        phone_book = contacts.all_contacts()
    except Exception:
        phone_book = {}
    wb = load_workbook(excel_path, read_only=True, data_only=True)
    ws = wb.active
    people = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or not row[0]:
            continue
        name = str(row[0]).strip()
        mobile = str(row[1] or "").strip().split(".")[0]  # 防 Excel 把手机号存成数字
        if not mobile:
            mobile = phone_book.get(name, "")  # 从 SQLite 联系人库补
        role = str(row[2] or "").strip()
        total = int(row[3] or 0)
        read = int(row[4] or 0)
        people.append(Person(name=name, mobile=mobile, role=role, total=total, read=read))
    wb.close()
    if not people:
        raise ValueError(f"{excel_path} 里没有数据行")
    return people


def make_template(excel_path: str):
    parent = os.path.dirname(excel_path)
    if parent:
        os.makedirs(parent, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "已读统计"
    ws.append(COLUMNS)
    # 示例数据，正式使用时删掉或改掉
    sample = [
        ["张三", "13800000001", "安全管理员", 7, 2],
        ["李四", "13800000002", "部门负责人", 30, 15],
        ["王五", "13800000003", "操作主管", 34, 25],
        ["赵六", "13800000004", "", 8, 8],
    ]
    for row in sample:
        ws.append(row)
    widths = [14, 16, 16, 10, 10]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[chr(64 + i)].width = w
    wb.save(excel_path)


# ---------- 数据日期 meta（跟随数据，不依赖推送时刻） ----------

def save_data_date(excel_path: str, date) -> None:
    """把数据所属日期写进 Excel 的 meta sheet，推送时读取，避免标题日期算错。"""
    wb = load_workbook(excel_path)
    if "meta" in wb.sheetnames:
        del wb["meta"]
    ws = wb.create_sheet("meta")
    ws["A1"] = "data_date"
    ws["B1"] = date.isoformat()
    wb.save(excel_path)


def load_data_date(excel_path: str):
    """读取数据日期；没有 meta 或读取失败返回 None。"""
    try:
        wb = load_workbook(excel_path, read_only=True, data_only=True)
        if "meta" not in wb.sheetnames:
            wb.close()
            return None
        row = next(wb["meta"].iter_rows(min_row=1, max_row=1, values_only=True))
        wb.close()
        import datetime as _dt
        return _dt.date.fromisoformat(str(row[1]).strip())
    except Exception:
        return None
