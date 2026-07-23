from __future__ import annotations

import re
import sys
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from services.excel_processor import LedgerData, LedgerProcessingError
from utils.paths import app_root


TARGET_AREAS = {
    "binhe": "滨河街道",
    "xinggu": "兴谷街道",
    "pinggu": "平谷镇",
}

ISSUE_DESCRIPTIONS = {
    "居民自主投放": "居民自主投放不规范",
    "容器满冒": "容器垃圾满冒堆积",
    "桶站及周边环境": "桶站及周边环境不整洁",
    "垃圾分类驿站": "垃圾分类驿站不规范",
    "容器整洁率": "容器脏污",
    "大件垃圾投放点设置": "大件垃圾投放点设置不规范",
    "容器标识颜色": "容器标识颜色不规范",
    "容器成组": "容器未成组摆放",
    "容器完好率": "容器破损",
    "桶站便利措施": "桶站便利措施不完善",
    "装修垃圾投放点设置": "装修垃圾投放点设置不规范",
}


@dataclass(frozen=True)
class RankedItem:
    name: str
    percent: str
    count: int = 0


class NoticeService:
    def generate_notice(self, problem: LedgerData, output_path: str | Path) -> Path:
        template = self._load_template()
        context = self._build_context(problem)
        text = self._format_notice(context) if self._is_default_template(template) else self._render(template, context)

        out_path = Path(output_path)
        if out_path.suffix.lower() != ".txt":
            out_path = out_path.with_suffix(".txt")
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_text(text, encoding="utf-8")
        return out_path

    def _load_template(self) -> str:
        template_path = self._template_path()
        if not template_path.exists():
            raise LedgerProcessingError(f"未找到通告模板：{template_path}")
        return template_path.read_text(encoding="utf-8")

    def _template_path(self) -> Path:
        names = ["input", "两街一镇通告模板.txt"]
        candidates = [app_root().joinpath(*names)]
        bundle_root = getattr(sys, "_MEIPASS", None)
        if bundle_root:
            candidates.append(Path(bundle_root).joinpath(*names))
        return next((path for path in candidates if path.exists()), candidates[0])

    def _build_context(self, problem: LedgerData) -> dict[str, Any]:
        required = ["3级点位", "5级点位", "2级指标"]
        missing = [name for name in required if name not in problem.headers]
        if missing:
            raise LedgerProcessingError("生成通告缺少必要字段：" + "、".join(missing))

        area_idx = problem.headers.index("3级点位")
        community_idx = problem.headers.index("5级点位")
        problem_type_idx = problem.headers.index("2级指标")

        context: dict[str, Any] = {"date": problem.date_label or "未识别日期"}
        for key, area_name in TARGET_AREAS.items():
            rows = [
                row for row in problem.rows
                if self._cell_text(row, area_idx) == area_name
            ]
            context[key] = self._summarize_area(rows, community_idx, problem_type_idx)
        return context

    def _summarize_area(
        self,
        rows: list[list[object]],
        community_idx: int,
        problem_type_idx: int,
    ) -> dict[str, Any]:
        problem_count = len(rows)
        community_names = [
            self._cell_text(row, community_idx)
            for row in rows
            if self._cell_text(row, community_idx)
        ]
        problem_types = [
            self._cell_text(row, problem_type_idx)
            for row in rows
            if self._cell_text(row, problem_type_idx)
        ]
        type_rank = self._rank(Counter(problem_types), problem_count, 2, describe_issue=True)
        location_rank = self._rank(Counter(community_names), problem_count, 3)
        return {
            "community_count": len(set(community_names)),
            "problem_count": problem_count,
            "top1": type_rank[0],
            "top2": type_rank[1],
            "location1": location_rank[0],
            "location2": location_rank[1],
            "location3": location_rank[2],
        }

    def _rank(
        self,
        counter: Counter[str],
        total: int,
        limit: int,
        describe_issue: bool = False,
    ) -> list[RankedItem]:
        first_seen = {name: index for index, name in enumerate(counter)}
        ranked = [
            RankedItem(self._display_issue_name(name) if describe_issue else name, self._percent(count, total), count)
            for name, count in sorted(counter.items(), key=lambda item: (-item[1], first_seen[item[0]]))
        ]
        ranked.extend(RankedItem("无", "0.0", 0) for _ in range(limit - len(ranked)))
        return ranked[:limit]

    def _percent(self, count: int, total: int) -> str:
        if total <= 0:
            return "0.0"
        value = round(count * 100 / total, 1)
        return f"{value:.1f}"

    def _display_issue_name(self, name: str) -> str:
        return ISSUE_DESCRIPTIONS.get(name, name)

    def _is_default_template(self, template: str) -> bool:
        return all(marker in template for marker in ("{{ binhe.location1.name }}", "{{ xinggu.top1.name }}", "{{ pinggu.top2.percent }}"))

    def _format_notice(self, context: dict[str, Any]) -> str:
        paragraphs = [
            self._format_area(context["date"], "滨河街道", context["binhe"], prefix=True),
            self._format_area("", "兴谷街道", context["xinggu"], prefix=False),
            self._format_area("", "平谷镇", context["pinggu"], prefix=False),
        ]
        return "\n".join(paragraphs)

    def _format_area(self, date_label: str, area_name: str, stats: dict[str, Any], prefix: bool) -> str:
        start = f"{date_label}，区级检查" if prefix else "检查"
        return (
            f"{start}{area_name}{stats['community_count']}个小区，发现问题{stats['problem_count']}处，"
            f"主要问题为{self._format_top_issues(stats)}。"
            f"其中{self._format_locations(stats)}。"
        )

    def _format_top_issues(self, stats: dict[str, Any]) -> str:
        items = [stats["top1"], stats["top2"]]
        return "、".join(f"{item.name}（{item.percent}%）" for item in items if item.count > 0)

    def _format_locations(self, stats: dict[str, Any]) -> str:
        items = [stats["location1"], stats["location2"], stats["location3"]]
        items = [item for item in items if item.count > 0]
        if not items:
            return "无小区问题"

        parts: list[str] = []
        index = 0
        while index < len(items):
            same_percent = [items[index]]
            index += 1
            while index < len(items) and items[index].percent == same_percent[0].percent:
                same_percent.append(items[index])
                index += 1

            names = self._join_names([item.name for item in same_percent])
            percent = same_percent[0].percent
            if len(same_percent) > 1:
                if not parts and len(same_percent) == len(items):
                    parts.append(f"{names}问题占比均为{percent}%")
                else:
                    parts.append(f"{names}问题占比均为{percent}%")
            elif not parts:
                parts.append(f"{names}问题率最高，占比{percent}%")
            else:
                parts.append(f"{names}问题占比{percent}%")
        return "，".join(parts)

    def _join_names(self, names: list[str]) -> str:
        if len(names) <= 1:
            return names[0] if names else ""
        if len(names) == 2:
            return "和".join(names)
        return "、".join(names[:-1]) + "和" + names[-1]

    def _render(self, template: str, context: dict[str, Any]) -> str:
        def replace(match: re.Match[str]) -> str:
            expression = match.group(1).strip()
            value: Any = context
            for part in expression.split("."):
                if isinstance(value, dict):
                    value = value.get(part, "")
                else:
                    value = getattr(value, part, "")
            return str(value)

        return re.sub(r"\{\{\s*([^{}]+?)\s*\}\}", replace, template)

    def _cell_text(self, row: list[object], index: int) -> str:
        if index >= len(row) or row[index] is None:
            return ""
        return str(row[index]).strip()
