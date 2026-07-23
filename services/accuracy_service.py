from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from services.excel_processor import LedgerData, LedgerProcessingError


CHECK_ERROR_RE = re.compile(r"查\s*(\d+)\s*错\s*(\d+)")
FULLWIDTH_DIGITS = str.maketrans("０１２３４５６７８９", "0123456789")


@dataclass
class AccuracyUpdate:
    names: list[str]
    error_count: int | str


class AccuracyService:
    def update_statistics(self, problem: LedgerData, statistics_path: str | Path) -> Path:
        source = Path(statistics_path)
        if source.suffix.lower() != ".xlsx":
            raise LedgerProcessingError("准确率统计表仅支持 .xlsx 文件")

        date_label = problem.date_label
        if not date_label:
            raise LedgerProcessingError("无法从问题台账文件名识别日期")
        date_value = problem.metadata.get("date_value")
        if not isinstance(date_value, date):
            date_value = None

        updates = self._collect_updates(problem)
        workbook = load_workbook(source)

        missing: list[str] = []
        for sheet_name, records in updates.items():
            if sheet_name not in workbook.sheetnames:
                raise LedgerProcessingError(f"统计表缺少 Sheet：{sheet_name}")
            sheet = workbook[sheet_name]
            name_col = self._find_name_col(sheet, sheet_name)
            error_col = self._find_accuracy_error_col(sheet)
            date_col = self._find_date_col(sheet)
            row_map = self._build_name_row_map(sheet, name_col)
            matched_updates: dict[int, int | str] = {}
            for record in records:
                row_index = self._match_location_row(row_map, record.names)
                if not row_index:
                    missing.append(f"{sheet_name}:{'/'.join(record.names)}")
                    continue
                if row_index not in matched_updates:
                    matched_updates[row_index] = record.error_count
                elif isinstance(record.error_count, int):
                    current = matched_updates[row_index]
                    matched_updates[row_index] = (current if isinstance(current, int) else 0) + record.error_count
            for row_index, error_count in matched_updates.items():
                sheet.cell(row=row_index, column=error_col).value = error_count
                date_cell = sheet.cell(row=row_index, column=date_col)
                date_cell.value = date_value if date_value else date_label
                if date_value:
                    date_cell.number_format = "m月d日"

        target = source.with_name(f"{date_label}小区村值守率及投放准确率统计.xlsx")
        workbook.save(target)
        if missing:
            raise LedgerProcessingError(
                "准确率统计表已保存，但以下点位未找到，未新增行：" + "、".join(missing[:20])
            )
        return target

    def _collect_updates(self, problem: LedgerData) -> dict[str, list[AccuracyUpdate]]:
        indexes = {name: self._header_index(problem.headers, name) for name in ("2级点位", "具体问题")}
        second_indicator_idx = self._optional_header_index(problem.headers, "2级指标")
        third_indicator_idx = self._optional_header_index(problem.headers, "3级指标")
        candidate_name_indexes = [
            problem.headers.index(name)
            for name in ("3级点位", "5级点位", "4级点位")
            if name in problem.headers
        ]
        if not candidate_name_indexes:
            raise LedgerProcessingError("缺少必要字段：3级点位")

        grouped: dict[str, dict[tuple[str, ...], dict[str, object]]] = {"小区": {}, "村居": {}}
        for row in problem.rows:
            point_type = self._cell_text(row, indexes["2级点位"])
            if point_type not in grouped:
                continue
            names = self._candidate_location_names(row, candidate_name_indexes)
            if not names:
                continue
            key = tuple(self._normalize_name(name) for name in names)
            record = grouped[point_type].setdefault(
                key,
                {
                    "names": names,
                    "check_error_sum": 0,
                    "has_check_error": False,
                    "has_fallback_match": False,
                    "fallback_count": 0,
                },
            )
            problem_text = self._cell_text(row, indexes["具体问题"])
            error_count = self._parse_error_count(problem_text)
            if error_count is not None:
                record["has_check_error"] = True
                record["check_error_sum"] = int(record["check_error_sum"]) + error_count

            second_indicator = self._cell_text(row, second_indicator_idx) if second_indicator_idx is not None else ""
            third_indicator = self._cell_text(row, third_indicator_idx) if third_indicator_idx is not None else ""
            is_resident_error = second_indicator == "居民自主投放" and third_indicator == "投放错误"
            is_station_problem = second_indicator == "垃圾分类驿站" and third_indicator != "无问题"
            if is_resident_error or is_station_problem:
                record["has_fallback_match"] = True
                record["fallback_count"] = int(record["fallback_count"]) + 1

        updates: dict[str, list[AccuracyUpdate]] = {"小区": [], "村居": []}
        for point_type, records in grouped.items():
            for record in records.values():
                error_value: int | str
                if record["has_check_error"]:
                    error_value = int(record["check_error_sum"])
                elif record["has_fallback_match"]:
                    error_value = int(record["fallback_count"])
                else:
                    error_value = "-"
                updates[point_type].append(AccuracyUpdate(list(record["names"]), error_value))
        return updates

    def _parse_error_count(self, text: str) -> int | None:
        normalized = (
            text.translate(FULLWIDTH_DIGITS)
            .replace("（", "(")
            .replace("）", ")")
            .replace("，", ",")
            .replace("：", ":")
        )
        match = CHECK_ERROR_RE.search(normalized)
        if not match:
            return None
        return int(match.group(2))

    def _candidate_location_names(self, row: list[object], indexes: list[int]) -> list[str]:
        result: list[str] = []
        for idx in indexes:
            value = self._cell_text(row, idx)
            if value and value not in {"小区", "村居", "点位"} and value not in result:
                result.append(value)
        return result

    def _match_location_row(self, row_map: dict[str, int], names: list[str]) -> int | None:
        for name in names:
            row_index = row_map.get(self._normalize_name(name))
            if row_index:
                return row_index
        return None

    def _find_name_col(self, sheet: Worksheet, sheet_name: str) -> int:
        candidates = ["小区名称"] if sheet_name == "小区" else ["村居名称", "村名称"]
        for row in sheet.iter_rows(min_row=1, max_row=min(sheet.max_row, 5)):
            for cell in row:
                if str(cell.value or "").strip() in candidates:
                    return cell.column
        raise LedgerProcessingError(f"{sheet_name} Sheet 未找到名称列")

    def _find_accuracy_error_col(self, sheet: Worksheet) -> int:
        header_row = self._find_header_row(sheet, "投放错误数")
        for cell in sheet[header_row]:
            if str(cell.value or "").strip() == "投放错误数":
                return cell.column
        raise LedgerProcessingError(f"{sheet.title} Sheet 未找到投放错误数列")

    def _find_date_col(self, sheet: Worksheet) -> int:
        for row in sheet.iter_rows(min_row=1, max_row=min(sheet.max_row, 5)):
            for cell in row:
                if str(cell.value or "").strip() == "日期":
                    return cell.column
        raise LedgerProcessingError(f"{sheet.title} Sheet 未找到日期列")

    def _find_header_row(self, sheet: Worksheet, text: str) -> int:
        for row in sheet.iter_rows(min_row=1, max_row=min(sheet.max_row, 5)):
            for cell in row:
                if str(cell.value or "").strip() == text:
                    return cell.row
        raise LedgerProcessingError(f"{sheet.title} Sheet 未找到字段：{text}")

    def _build_name_row_map(self, sheet: Worksheet, name_col: int) -> dict[str, int]:
        result: dict[str, int] = {}
        for row_index in range(4, sheet.max_row + 1):
            name = self._normalize_name(sheet.cell(row=row_index, column=name_col).value)
            if name and name not in result:
                result[name] = row_index
        return result

    def _header_index(self, headers: list[str], name: str) -> int:
        try:
            return headers.index(name)
        except ValueError as exc:
            raise LedgerProcessingError(f"缺少必要字段：{name}") from exc

    def _optional_header_index(self, headers: list[str], name: str) -> int | None:
        try:
            return headers.index(name)
        except ValueError:
            return None

    def _cell_text(self, row: list[object], index: int) -> str:
        if index >= len(row) or row[index] is None:
            return ""
        return str(row[index]).strip()

    def _normalize_name(self, value: object) -> str:
        return re.sub(r"\s+", "", str(value or "")).strip()
