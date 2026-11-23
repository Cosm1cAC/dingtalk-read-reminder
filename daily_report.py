# -*- coding: utf-8 -*-
"""每日助手数据通报：自动发现每日报表目录中的最新导出表，计算三项重点指标并组装四块通报内容。

三项重点指标：
  1. 中心管理助手-使用率（数据源：部长助手-员工明细，日数据，按已读占比 >= 50% 人数 / 总人数）
  2. 中心管理助手-问题解决率（总部周报参考值，配置于 config.yaml report.national_weekly.manager_solve）
  3. 操作工助手-改善率（数据源：操作工改善报表-明细，周数据；辅以操作异常培训人员名单给出绩差基准差距）

通报四块（缺一不可）：
  ① 整体结论（全国及本省区三项重点指标数值，明确标注日/周口径）
  ② 末位部门（本省区排名末3位：部门名称、人数、指标数值）
  ③ 末位人员（排名末5位：姓名、所属部门、指标数值、与基础值的差距）
  ④ 整改要求（@到具体人、今日动作、复查时间）
"""
from collections import defaultdict
from dataclasses import dataclass, field
import datetime
import os
import re
import warnings

from openpyxl import load_workbook

import contacts

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")


@dataclass
class DeptStat:
    center: str
    name: str
    total: int
    achieved: int
    unimproved_names: list = field(default_factory=list)

    @property
    def rate(self) -> float:
        return (self.achieved / self.total * 100.0) if self.total else 0.0


@dataclass
class ManagerPersonStat:
    name: str
    center: str
    role: str
    read_cnt: int
    msg_cnt: int
    rate: float  # 百分比 0~100
    emp_id: str = ""
    dept_name: str = ""

    @property
    def dept(self) -> str:
        return f"{self.center}-{self.dept_name}" if self.dept_name else self.center

    @property
    def gap_to_base(self) -> float:
        return self.rate - 50.0


@dataclass
class OperatorAbnormalStat:
    name: str
    center: str
    dept_name: str
    metric_name: str
    value: float
    base: float
    is_abnormal_metric: bool = True  # False 时退化为展示异常次数

    @property
    def dept(self) -> str:
        return f"{self.center}-{self.dept_name}" if self.dept_name else self.center

    @property
    def gap(self) -> float:
        return self.value - self.base

    @property
    def pct_over(self) -> float:
        return (self.gap / self.base * 100.0) if self.base else 0.0


@dataclass
class DailyReportData:
    province: str
    daily_date: datetime.date
    improve_date: datetime.date
    improve_period: str = ""
    dept_level: str = "team"
    min_people_improve: int = 3

    # 1. 中心管理助手-使用率
    mgr_total: int = 0
    mgr_used: int = 0
    mgr_rate: float = 0.0
    mgr_month_rate: float = None
    mgr_centers: list = field(default_factory=list)
    mgr_bottom_people: list = field(default_factory=list)
    mgr_unused_count: int = 0

    # 2. 操作工助手-改善率
    imp_nat_total: int = 0
    imp_nat_ok: int = 0
    imp_nat_rate: float = 0.0
    imp_prov_total: int = 0
    imp_prov_ok: int = 0
    imp_prov_rate: float = 0.0
    imp_prov_train_rate: float = 0.0
    imp_prov_rank: int = 0
    imp_total_provinces: int = 0
    imp_centers: list = field(default_factory=list)
    imp_bottom_depts: list = field(default_factory=list)
    imp_bottom_people: list = field(default_factory=list)
    imp_has_perf_detail: bool = False

    # 班组负责人查找表（来自推送消息汇总，绝不含部长/中心负责人）
    team_leaders_exact: dict = field(default_factory=dict)
    team_leaders_area: dict = field(default_factory=dict)

    source_files: dict = field(default_factory=dict)


def _find_latest_file(report_dir: str, keyword: str, required: bool = True) -> str:
    if not os.path.isdir(report_dir):
        if required:
            raise FileNotFoundError(f"找不到报表目录: {report_dir}")
        return ""
    candidates = [
        fn for fn in os.listdir(report_dir)
        if fn.endswith(".xlsx") and not fn.startswith("~$") and keyword in fn
    ]
    if not candidates:
        if required:
            raise FileNotFoundError(f"在 {report_dir} 中未找到包含「{keyword}」的 Excel 文件")
        return ""
    # BI 导出文件名通常带 _YYYYMMDDHHMMSS 后缀，按文件名降序优先，辅以修改时间
    candidates.sort(
        key=lambda fn: (fn, os.path.getmtime(os.path.join(report_dir, fn))),
        reverse=True,
    )
    return os.path.join(report_dir, candidates[0])


def _read_sheet_rows(path: str) -> list:
    """读取首张工作表为 dict 列表。重置 dimension 避免 BI 导出元信息截断行数。"""
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb.active
        ws.reset_dimensions()
        it = ws.iter_rows(values_only=True)
        raw_headers = next(it, None)
        if not raw_headers:
            raise ValueError(f"{path} 为空表")
        headers = [str(h).strip() if h is not None else "" for h in raw_headers]
        rows = []
        for raw in it:
            if not raw or all(v is None or str(v).strip() == "" for v in raw):
                continue
            row_dict = {headers[i]: raw[i] if i < len(raw) else None
                        for i in range(len(headers)) if headers[i]}
            rows.append(row_dict)
        return rows
    finally:
        wb.close()


def _to_float(val, default: float = 0.0) -> float:
    if val is None or str(val).strip() == "":
        return default
    return float(val)


def _to_int(val, default: int = 0) -> int:
    return int(round(_to_float(val, float(default))))


def _to_str(val) -> str:
    return "" if val is None else str(val).strip()


def _parse_date(val) -> datetime.date:
    if isinstance(val, datetime.datetime):
        return val.date()
    if isinstance(val, datetime.date):
        return val
    return datetime.date.fromisoformat(str(val).strip()[:10])


def _dept_key(dept_name: str, dept_level: str) -> str:
    name = _to_str(dept_name)
    if dept_level == "area" and "_" in name:
        return name.split("_", 1)[0]
    return name or "未分组"


def _group_bottom_depts(rows: list, center_field: str, dept_field: str,
                        ok_fn, dept_level: str, min_people: int,
                        limit: int = 3, name_field: str = "") -> tuple:
    """同时返回全部中心汇总与本省区末 N 个部门（可选收集未达标姓名）。"""
    by_center = defaultdict(lambda: [0, 0])
    by_dept = defaultdict(lambda: [0, 0])
    unimproved_by_dept = defaultdict(list)
    seen_unimproved = defaultdict(set)

    for r in rows:
        center = _to_str(r.get(center_field)) or "未知中心"
        dept = _dept_key(r.get(dept_field), dept_level)
        ok = 1 if ok_fn(r) else 0
        by_center[center][0] += 1
        by_center[center][1] += ok
        by_dept[(center, dept)][0] += 1
        by_dept[(center, dept)][1] += ok
        if not ok and name_field:
            emp_name = _to_str(r.get(name_field))
            if emp_name and emp_name not in seen_unimproved[(center, dept)]:
                seen_unimproved[(center, dept)].add(emp_name)
                unimproved_by_dept[(center, dept)].append(emp_name)

    centers = [DeptStat(center=c, name=c, total=v[0], achieved=v[1])
               for c, v in by_center.items()]
    centers.sort(key=lambda d: (d.rate, -d.total, d.name))

    all_depts = [
        DeptStat(
            center=c,
            name=f"{c}-{d}",
            total=v[0],
            achieved=v[1],
            unimproved_names=unimproved_by_dept.get((c, d), []),
        )
        for (c, d), v in by_dept.items()
    ]
    qualified = [d for d in all_depts if d.total >= min_people]
    pool = qualified if len(qualified) >= limit else all_depts
    pool.sort(key=lambda d: (d.rate, -d.total, d.name))
    return centers, pool[:limit]


def _parse_manager_usage(report_dir: str, province: str) -> dict:
    detail_path = _find_latest_file(report_dir, "部长助手-员工明细", required=True)
    rows = _read_sheet_rows(detail_path)
    day_rows = [
        r for r in rows
        if _to_str(r.get("日期类型")) == "日"
        and (not province or not _to_str(r.get("省区名称")) or _to_str(r.get("省区名称")) == province)
    ]
    if not day_rows:
        raise ValueError(f"{detail_path} 中没有「日」类型且省区为 {province} 的数据")

    max_date = max(_parse_date(r["日期"]) for r in day_rows)
    day_rows = [r for r in day_rows if _parse_date(r["日期"]) == max_date]

    by_center = defaultdict(lambda: [0, 0])
    people = []
    for r in day_rows:
        center = _to_str(r.get("中心名称")) or "未知中心"
        rate = _to_float(r.get("已读占比(%)"))
        is_used = rate >= 50.0
        by_center[center][0] += 1
        by_center[center][1] += 1 if is_used else 0
        people.append(ManagerPersonStat(
            name=_to_str(r.get("员工名称")),
            center=center,
            role=_to_str(r.get("岗位名称")),
            read_cnt=_to_int(r.get("已读消息量")),
            msg_cnt=_to_int(r.get("消息量")),
            rate=rate,
            emp_id=_to_str(r.get("员工工号")),
        ))

    centers = [DeptStat(center=c, name=c, total=v[0], achieved=v[1])
               for c, v in by_center.items()]
    centers.sort(key=lambda d: (d.rate, -d.total, d.name))

    people.sort(key=lambda p: (p.rate, -p.msg_cnt, p.name))
    total = len(people)
    used = sum(1 for p in people if p.rate >= 50.0)

    # 可选读取部长助手-汇总表的当月累计使用率
    month_rate = None
    sum_path = _find_latest_file(report_dir, "部长助手-汇总表", required=False)
    if sum_path:
        for r in _read_sheet_rows(sum_path):
            if (_to_str(r.get("日期类型")) == "月"
                    and _to_str(r.get("组织类型")) == "省区"
                    and _to_str(r.get("职位类型")) == "全部"
                    and (not province or _to_str(r.get("省区名称")) == province)):
                month_rate = _to_float(r.get("使用率(%)"))
                break

    return {
        "date": max_date,
        "total": total,
        "used": used,
        "rate": (used / total * 100.0) if total else 0.0,
        "month_rate": month_rate,
        "centers": centers,
        "bottom_people": people[:5],
        "unused_count": sum(1 for p in people if p.rate < 50.0),
        "file": detail_path,
    }


def _parse_operator_improve(report_dir: str, province: str,
                            dept_level: str, min_people: int) -> dict:
    imp_path = _find_latest_file(report_dir, "操作工改善报表-明细", required=True)
    rows = _read_sheet_rows(imp_path)
    max_date = max(_parse_date(r["日期"]) for r in rows)
    rows = [r for r in rows if _parse_date(r["日期"]) == max_date]

    nat_total = len(rows)
    nat_ok = sum(1 for r in rows if _to_str(r.get("是否改善")) == "是")

    by_prov = defaultdict(lambda: [0, 0])
    for r in rows:
        p = _to_str(r.get("省区名称")) or "未知省区"
        by_prov[p][0] += 1
        by_prov[p][1] += 1 if _to_str(r.get("是否改善")) == "是" else 0
    prov_ranking = sorted(
        ((v[1] / v[0] * 100.0 if v[0] else 0.0, v[1], p) for p, v in by_prov.items()),
        key=lambda x: (-x[0], -x[1], x[2]),
    )
    prov_rank = next((i + 1 for i, item in enumerate(prov_ranking) if item[2] == province), 0)

    prov_rows = [r for r in rows if _to_str(r.get("省区名称")) == province]
    if not prov_rows:
        raise ValueError(f"{imp_path} 中没有省区为 {province} 的数据")

    prov_total = len(prov_rows)
    prov_ok = sum(1 for r in prov_rows if _to_str(r.get("是否改善")) == "是")
    prov_train_ok = sum(1 for r in prov_rows if _to_str(r.get("是否完成培训")) == "是")

    centers, bottom_depts = _group_bottom_depts(
        prov_rows, "中心名称", "班组名称",
        lambda r: _to_str(r.get("是否改善")) == "是",
        dept_level, min_people, limit=3, name_field="员工名称",
    )

    # 结合操作异常培训人员名单给出个人指标与基准值差距
    abn_path = _find_latest_file(report_dir, "操作异常培训人员名单", required=False)
    period_str = ""
    bottom_people = []
    has_perf_detail = False

    if abn_path:
        abn_rows = _read_sheet_rows(abn_path)
        prov_centers = {_to_str(r.get("中心名称")) for r in prov_rows if _to_str(r.get("中心名称"))}
        abn_sc = [r for r in abn_rows if _to_str(r.get("中心")) in prov_centers]
        if abn_sc:
            abn_max_date = max(_parse_date(r["数据日期"]) for r in abn_sc)
            abn_sc = [r for r in abn_sc if _parse_date(r["数据日期"]) == abn_max_date]
            period_str = _to_str(abn_sc[0].get("时间周期"))
            imp_map = {_to_str(r.get("员工编码")): r for r in prov_rows if _to_str(r.get("员工编码"))}

            # 按工号保留相对基准值超标比例最高的一条记录
            worst_unimproved = {}
            worst_other = {}
            for r in abn_sc:
                eid = _to_str(r.get("工号"))
                val = _to_float(r.get("指标值"))
                base = _to_float(r.get("绩效考核基准值"))
                if not eid or base <= 0 or val <= base:
                    continue
                ratio = (val - base) / base
                imp_rec = imp_map.get(eid)
                dept_name = _dept_key(imp_rec.get("班组名称"), dept_level) if imp_rec else ""
                stat = OperatorAbnormalStat(
                    name=_to_str(r.get("姓名")),
                    center=_to_str(r.get("中心")),
                    dept_name=dept_name,
                    metric_name=_to_str(r.get("指标中文名")),
                    value=val,
                    base=base,
                    is_abnormal_metric=True,
                )
                target_bucket = worst_unimproved if (imp_rec and _to_str(imp_rec.get("是否改善")) == "否") else worst_other
                prev = target_bucket.get(eid)
                if prev is None or (ratio, stat.gap) > (prev[0], prev[1].gap):
                    target_bucket[eid] = (ratio, stat)

            picked = sorted(worst_unimproved.values(),
                            key=lambda x: (-x[0], -x[1].gap, x[1].name))
            if len(picked) < 5:
                existing_ids = set(worst_unimproved.keys())
                extra = sorted(
                    ((eid, v) for eid, v in worst_other.items() if eid not in existing_ids),
                    key=lambda item: (-item[1][0], -item[1][1].gap, item[1][1].name),
                )
                picked.extend(v for _, v in extra)
            bottom_people = [item[1] for item in picked[:5]]
            has_perf_detail = bool(bottom_people)

    if not bottom_people:
        # 无绩差明细表时退化为改善表内未改善人员按12周异常次数排序
        unimproved = [r for r in prov_rows if _to_str(r.get("是否改善")) == "否"] or prov_rows
        unimproved.sort(
            key=lambda r: (-_to_float(r.get("最近12周异常次数")),
                           -_to_float(r.get("最近2周异常次数")),
                           _to_str(r.get("员工名称"))),
        )
        bottom_people = [
            OperatorAbnormalStat(
                name=_to_str(r.get("员工名称")),
                center=_to_str(r.get("中心名称")),
                dept_name=_dept_key(r.get("班组名称"), dept_level),
                metric_name="近12周异常次数",
                value=_to_float(r.get("最近12周异常次数")),
                base=0.0,
                is_abnormal_metric=False,
            )
            for r in unimproved[:5]
        ]

    return {
        "date": max_date,
        "period": period_str,
        "nat_total": nat_total,
        "nat_ok": nat_ok,
        "nat_rate": (nat_ok / nat_total * 100.0) if nat_total else 0.0,
        "prov_total": prov_total,
        "prov_ok": prov_ok,
        "prov_rate": (prov_ok / prov_total * 100.0) if prov_total else 0.0,
        "prov_train_rate": (prov_train_ok / prov_total * 100.0) if prov_total else 0.0,
        "prov_rank": prov_rank,
        "total_provinces": len(by_prov),
        "centers": centers,
        "bottom_depts": bottom_depts,
        "bottom_people": bottom_people,
        "has_perf_detail": has_perf_detail,
        "imp_file": imp_path,
        "abn_file": abn_path,
    }


def _load_team_leaders(report_dir: str) -> tuple:
    """从推送消息汇总中提取各班组组长/主管（排除中心部长）及员工工号->部门映射。"""
    path = _find_latest_file(report_dir, "操作工助手_推送消息汇总", required=False)
    if not path:
        return {}, {}, {}, ""

    exact_map = defaultdict(list)
    area_map = defaultdict(list)
    emp_depts = {}
    seen_exact = set()
    seen_area = set()
    for r in _read_sheet_rows(path):
        center = _to_str(r.get("组织名称")).replace("转运中心", "")
        dept = _to_str(r.get("部门"))
        role = _to_str(r.get("岗位"))
        name = _to_str(r.get("姓名"))
        emp_id = _to_str(r.get("工号"))
        if emp_id and dept:
            emp_depts.setdefault(emp_id, dept)
        if not center or not dept or not name:
            continue
        if "部长" in role or "高级经理" in role:
            continue
        if "组长" in role or "主管" in role:
            key = (center, dept)
            if (key, name) not in seen_exact:
                seen_exact.add((key, name))
                exact_map[key].append((name, role))
        if "主管" in role:
            key = (center, dept)
            if (key, name) not in seen_area:
                seen_area.add((key, name))
                area_map[key].append((name, role))

    for key in exact_map:
        exact_map[key].sort(key=lambda item: (0 if "组长" in item[1] else 1, item[0]))
    return dict(exact_map), dict(area_map), emp_depts, os.path.basename(path)


def load_daily_report(cfg: dict) -> DailyReportData:
    """从配置指定的每日报表目录自动加载并计算全部通报数据。"""
    rcfg = cfg.get("report") or {}
    report_dir = rcfg.get("report_dir") or "每日报表"
    province = (rcfg.get("province") or "").strip()
    if not province:
        raise ValueError("请在 config.yaml 的 report.province 中填写本省区名称")
    dept_level = rcfg.get("dept_level") or "team"
    min_improve = int(rcfg.get("min_people_improve", 3))

    mgr = _parse_manager_usage(report_dir, province)
    imp = _parse_operator_improve(report_dir, province, dept_level, min_improve)
    exact_leaders, area_leaders, emp_depts, leader_src = _load_team_leaders(report_dir)
    for p in mgr["bottom_people"]:
        if p.emp_id and p.emp_id in emp_depts:
            p.dept_name = emp_depts[p.emp_id]

    sources = {
        "manager_detail": os.path.basename(mgr["file"]),
        "operator_improve": os.path.basename(imp["imp_file"]),
    }
    if imp["abn_file"]:
        sources["operator_abnormal"] = os.path.basename(imp["abn_file"])
    if leader_src:
        sources["team_leaders_source"] = leader_src

    return DailyReportData(
        province=province,
        daily_date=mgr["date"],
        improve_date=imp["date"],
        improve_period=imp["period"],
        dept_level=dept_level,
        min_people_improve=min_improve,
        mgr_total=mgr["total"],
        mgr_used=mgr["used"],
        mgr_rate=mgr["rate"],
        mgr_month_rate=mgr["month_rate"],
        mgr_centers=mgr["centers"],
        mgr_bottom_people=mgr["bottom_people"],
        mgr_unused_count=mgr["unused_count"],
        imp_nat_total=imp["nat_total"],
        imp_nat_ok=imp["nat_ok"],
        imp_nat_rate=imp["nat_rate"],
        imp_prov_total=imp["prov_total"],
        imp_prov_ok=imp["prov_ok"],
        imp_prov_rate=imp["prov_rate"],
        imp_prov_train_rate=imp["prov_train_rate"],
        imp_prov_rank=imp["prov_rank"],
        imp_total_provinces=imp["total_provinces"],
        imp_centers=imp["centers"],
        imp_bottom_depts=imp["bottom_depts"],
        imp_bottom_people=imp["bottom_people"],
        imp_has_perf_detail=imp["has_perf_detail"],
        team_leaders_exact=exact_leaders,
        team_leaders_area=area_leaders,
        source_files=sources,
    )


def _fmt_num(x: float) -> str:
    if abs(x - round(x)) < 1e-6:
        return str(int(round(x)))
    return f"{x:.2f}".rstrip("0").rstrip(".")


def _fmt_signed(x: float, suffix: str = "%") -> str:
    sign = "+" if x >= 0 else "-"
    return f"{sign}{_fmt_num(abs(x))}{suffix}"


def _short_province(province: str) -> str:
    return province.replace("省区", "") if province.endswith("省区") else province


def _fmt_date_cn(d: datetime.date) -> str:
    return f"{d.month}月{d.day}日"


def _status_icon(rate: float, baseline: float, target: float, in_bottom5: bool = False) -> str:
    if rate < baseline or in_bottom5:
        return "🔴"
    if rate >= target:
        return "✅"
    return "⚠️"


def _fmt_period(period: str, fallback: str) -> str:
    m = re.match(r"\d{4}-(\d{2}-\d{2}).*?~\s*\d{4}-(\d{2}-\d{2})", _to_str(period))
    return f"{m.group(1)} ~ {m.group(2)}" if m else fallback


def mask_mobiles(text: str) -> str:
    """将文本中的 11 位手机号脱敏为 138****0001，仅用于终端日志展示。"""
    return re.sub(r"(?<!\d)(1\d{2})\d{4}(\d{4})(?!\d)", r"\1****\2", text)


def render_at_names(text: str, mobile_to_name: dict) -> str:
    """将 @手机号 还原为 @姓名，仅用于终端预览展示实际钉钉显示效果。"""
    return re.sub(
        r"@(1\d{9,10})(?!\d)",
        lambda m: f"@{mobile_to_name[m.group(1)]}" if m.group(1) in mobile_to_name else m.group(0),
        text,
    )


def build_daily_messages(data: DailyReportData, cfg: dict) -> dict:
    """组装钉钉双消息通报：主统计卡片 (actionCard) + 整改责任人 @提醒 (markdown)。

    返回 dict:
        title, card_text, at_title, at_text, at_mobiles, mobile_to_name,
        missing_phone_names, unmatched_teams
    """
    rcfg = cfg.get("report") or {}
    nw = rcfg.get("national_weekly") or {}
    weekly_as_of = nw.get("as_of") or "最新周报"
    mgr_cfg = nw.get("manager_usage") or {}
    solve_cfg = nw.get("manager_solve") or {}
    imp_cfg = nw.get("operator_improve") or {}

    mgr_base = float(mgr_cfg.get("baseline", 80.0))
    mgr_target = float(mgr_cfg.get("target", 85.0))
    solve_base = float(solve_cfg.get("baseline", 65.0))
    solve_target = float(solve_cfg.get("target", 70.0))
    imp_base = float(imp_cfg.get("baseline", 55.0))
    imp_target = float(imp_cfg.get("target", 60.0))

    prov_short = _short_province(data.province)
    daily_cn = _fmt_date_cn(data.daily_date)
    imp_cn = _fmt_date_cn(data.improve_date)
    title = f"📊 {daily_cn} {prov_short}数字化助手日报"

    action_text = rcfg.get("rectification_action") or (
        "请以上未读人员今日 18:00 前完成中心管理助手全部消息查看（严禁一键已读）；"
        "请末位班组负责人今日内组织未改善绩差人员一对一复盘培训，督促改善"
    )

    # ---------- 消息 1：主通报卡片（actionCard，三项重点指标 + 四块结构，不含任何职位） ----------
    mgr_rank = mgr_cfg.get("province_rank")
    mgr_total_p = mgr_cfg.get("total_provinces", 32)
    mgr_icon = _status_icon(
        data.mgr_rate, mgr_base, mgr_target,
        in_bottom5=bool(mgr_rank and mgr_total_p and int(mgr_rank) > int(mgr_total_p) - 5),
    )
    mgr_month_str = f"，本月累计 {data.mgr_month_rate:.2f}%" if data.mgr_month_rate is not None else ""
    nat_mgr = mgr_cfg.get("national")
    prov_mgr_wk = mgr_cfg.get("province_weekly")
    if nat_mgr is not None:
        wk_sub = []
        if prov_mgr_wk is not None:
            wk_sub.append(f"{prov_short}周报 {float(prov_mgr_wk):.2f}%")
        if mgr_rank is not None:
            wk_sub.append(f"排名 **第{int(mgr_rank)}/{int(mgr_total_p)}名**" if mgr_total_p else f"排名 **第{int(mgr_rank)}名**")
        mgr_nat_line = f"• 全国周报 **{float(nat_mgr):.2f}%**" + (f"｜{'，'.join(wk_sub)}" if wk_sub else "")
    else:
        mgr_nat_line = "• 全国周报：总部周报暂无，待补充"

    solve_nat = solve_cfg.get("national")
    solve_prov = solve_cfg.get("province_weekly")
    if solve_nat is not None and solve_prov is not None:
        solve_nat_f = float(solve_nat)
        solve_prov_f = float(solve_prov)
        solve_rank = solve_cfg.get("province_rank")
        solve_total_p = solve_cfg.get("total_provinces", 32)
        solve_icon = _status_icon(
            solve_prov_f, solve_base, solve_target,
            in_bottom5=bool(solve_rank and solve_total_p and int(solve_rank) > int(solve_total_p) - 5),
        )
        solve_rank_str = f"排名 **第{int(solve_rank)}/{int(solve_total_p)}名**" if solve_rank is not None else "排名待补充"
        solve_lines = [
            f"**2. 中心管理助手 · 问题解决率** {solve_icon}（总部周报）",
            f"• {prov_short} **{solve_prov_f:.2f}%**｜基础 {solve_base:.0f}%（{solve_prov_f - solve_base:+.2f}%）/ 目标 {solve_target:.0f}%",
            f"• 全国 **{solve_nat_f:.2f}%**｜{prov_short}{solve_rank_str}（较全国 {solve_prov_f - solve_nat_f:+.2f}%）",
        ]
    else:
        solve_lines = [
            "**2. 中心管理助手 · 问题解决率** ⚠️（总部周报）",
            "• 总部周报暂无，待补充",
        ]

    from_bottom = (data.imp_total_provinces - data.imp_prov_rank + 1) if data.imp_prov_rank else 0
    imp_icon = _status_icon(
        data.imp_prov_rate, imp_base, imp_target,
        in_bottom5=bool(data.imp_prov_rank and data.imp_total_provinces and data.imp_prov_rank > data.imp_total_provinces - 5),
    )
    imp_rank_str = (
        f"{prov_short}排名 **第{data.imp_prov_rank}/{data.imp_total_provinces}名（倒数第{from_bottom}）**"
        if data.imp_prov_rank else "未匹配到全国排名"
    )

    card_lines = [
        f"### {title}",
        (
            f"> **口径说明**：中心管理使用率为昨日（{daily_cn}）日数据；操作工改善率为最新周数据（截止{imp_cn}）；"
            f"全国数值、全国排名及问题解决率取自总部周报（{weekly_as_of}）。"
        ),
        "---",
        f"#### 📈 ① 整体结论（全国 vs {prov_short}）",
        f"**1. 中心管理助手 · 使用率** {mgr_icon}",
        f"• {prov_short}昨日 **{data.mgr_rate:.2f}%**（{data.mgr_used}/{data.mgr_total}人{mgr_month_str}）｜基础 {mgr_base:.0f}%（{data.mgr_rate - mgr_base:+.2f}%）/ 目标 {mgr_target:.0f}%",
        mgr_nat_line,
        *solve_lines,
        f"**3. 操作工助手 · 改善率** {imp_icon}（周截止{imp_cn}）",
        f"• {prov_short} **{data.imp_prov_rate:.2f}%**（{data.imp_prov_ok}/{data.imp_prov_total}人）｜基础 {imp_base:.0f}%（{data.imp_prov_rate - imp_base:+.2f}%）/ 目标 {imp_target:.0f}%",
        f"• 全国 **{data.imp_nat_rate:.2f}%**（{data.imp_nat_ok}/{data.imp_nat_total}人）｜{imp_rank_str}，较全国 {data.imp_prov_rate - data.imp_nat_rate:+.2f}%",
        "---",
        f"#### 📉 ② 末位部门（{prov_short}排名末3位）",
        "**▸ 中心管理助手 · 使用率（按中心）**",
    ]

    for i, d in enumerate(data.mgr_centers[:3], 1):
        c_icon = _status_icon(d.rate, mgr_base, mgr_target)
        card_lines.append(
            f"{i}. **{d.name}**｜{d.total}人｜**{d.rate:.2f}%**（{d.achieved}/{d.total}，较基础值 {d.rate - mgr_base:+.2f}%）{c_icon}"
        )

    card_lines.append(f"**▸ 操作工助手 · 改善率（≥{data.min_people_improve}人班组末3）**")
    for i, d in enumerate(data.imp_bottom_depts, 1):
        names_str = "、".join(d.unimproved_names) if d.unimproved_names else "无"
        card_lines.append(f"{i}. **{d.name}**｜{d.total}人｜改善率 **{d.rate:.2f}%**")
        card_lines.append(f"> 未改善名单：{names_str}")

    card_lines.extend([
        "---",
        "#### ⛔ ③ 末位人员（排名末5位）",
        f"**▸ 中心管理助手 · 使用率末5人**（基础值 50%，昨日共 {data.mgr_unused_count} 人未达标）",
    ])
    for i, p in enumerate(data.mgr_bottom_people, 1):
        card_lines.append(
            f"{i}. **{p.name}**｜{p.dept}｜**{p.rate:.2f}%**（{p.read_cnt}/{p.msg_cnt}）｜距基础值 **{p.gap_to_base:+.2f}%**"
        )

    if data.imp_has_perf_detail:
        period_label = _fmt_period(data.improve_period, imp_cn)
        card_lines.append(f"**▸ 操作工助手 · 绩差未改善末5人**（{period_label}）")
        for i, p in enumerate(data.imp_bottom_people, 1):
            card_lines.append(
                f"{i}. **{p.name}**｜{p.dept}｜{p.metric_name} **{_fmt_num(p.value)}**｜超基准{_fmt_num(p.base)} **{_fmt_signed(p.gap, '')}**"
            )
    else:
        card_lines.append(f"**▸ 操作工助手 · 未改善末5人**（周截止{imp_cn}）")
        for i, p in enumerate(data.imp_bottom_people, 1):
            card_lines.append(
                f"{i}. **{p.name}**｜{p.dept}｜近12周异常 **{_fmt_num(p.value)}次**"
            )

    card_lines.extend([
        "---",
        "📌 具体责任人 @提醒与今日整改要求见下条消息；如有优化建议，欢迎随时提出！",
    ])

    # ---------- 消息 2：责任人 @提醒（markdown，专做块④与@人，不重复卡片内容与重名） ----------
    phone_book = contacts.all_contacts()
    at_mobiles = []
    seen_mobiles = set()
    mobile_to_name = {}
    missing_phone_names = []

    def resolve_at(name: str) -> str:
        mobile = phone_book.get(name, "")
        if mobile:
            mobile_to_name[mobile] = name
            if mobile not in seen_mobiles:
                seen_mobiles.add(mobile)
                at_mobiles.append(mobile)
            return f"@{mobile}"
        if name not in missing_phone_names:
            missing_phone_names.append(name)
        return f"@{name}"

    at_title = f"🔔 {daily_cn} ④ 整改要求与责任人 @提醒"
    mgr_mentions = [resolve_at(p.name) for p in data.mgr_bottom_people]
    at_lines = [
        f"### {at_title}",
        "**⛔ 中心管理使用率末5人**",
        f"• {' '.join(mgr_mentions)}",
    ]

    team_owners = rcfg.get("team_owners") or {}

    def _manual_leaders(key: str) -> list:
        val = team_owners.get(key)
        if not val:
            return []
        items = [val] if isinstance(val, str) else list(val)
        return [(_to_str(x), "") for x in items if _to_str(x)]

    def _find_team_leaders(dept: DeptStat) -> list:
        team_key = dept.name
        raw_team = dept.name[len(dept.center) + 1:] if dept.name.startswith(dept.center + "-") else dept.name
        raw_area = raw_team.split("_", 1)[0]
        area_key = f"{dept.center}-{raw_area}"
        return (
            _manual_leaders(team_key)
            or _manual_leaders(raw_team)
            or data.team_leaders_exact.get((dept.center, raw_team), [])
            or _manual_leaders(area_key)
            or _manual_leaders(raw_area)
            or data.team_leaders_area.get((dept.center, raw_area), [])
        )

    unmatched_teams = []
    if data.imp_bottom_depts:
        at_lines.append("---")
        at_lines.append("**⚠️ 操作工改善末位班组**")
        for d in data.imp_bottom_depts:
            leaders = _find_team_leaders(d)
            if leaders:
                mentions = [resolve_at(name) for name, _ in leaders]
                at_lines.append(f"• {' '.join(mentions)}｜{d.name}")
            else:
                unmatched_teams.append(d.name)
                at_lines.append(f"• {d.name}（暂未匹配到班组负责人）")

    at_lines.extend([
        "---",
        f"• **今日动作**：{action_text}",
    ])

    # 钉钉 markdown 单个 \n 会被折叠，段与段之间使用空行 \n\n
    return {
        "title": title,
        "card_text": "\n\n".join(card_lines),
        "at_title": at_title,
        "at_text": "\n\n".join(at_lines),
        "at_mobiles": at_mobiles,
        "mobile_to_name": mobile_to_name,
        "missing_phone_names": missing_phone_names,
        "unmatched_teams": unmatched_teams,
    }


def build_daily_briefing(data: DailyReportData, cfg: dict) -> tuple:
    """兼容接口：返回拼接后的单文本及 @名单。"""
    msg = build_daily_messages(data, cfg)
    combined = msg["card_text"] + "\n\n---\n\n" + msg["at_text"]
    return msg["title"], combined, msg["at_mobiles"], msg["missing_phone_names"]
