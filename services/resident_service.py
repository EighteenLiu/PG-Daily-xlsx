from __future__ import annotations

from services.excel_processor import LedgerData, LedgerImage, LedgerProcessingError


class ResidentService:
    def make_resident_ledger(self, data: LedgerData) -> LedgerData:
        try:
            indicator_idx = data.headers.index("2级指标")
        except ValueError as exc:
            raise LedgerProcessingError("缺少必要字段：2级指标") from exc

        rows: list[list[object]] = []
        row_map: dict[int, int] = {}
        for old_index, row in enumerate(data.rows):
            value = row[indicator_idx] if indicator_idx < len(row) else None
            if str(value or "").strip() != "居民自主投放":
                continue
            row_map[old_index] = len(rows)
            rows.append(row[:])

        self._renumber_rows(data.headers, rows)

        images = [
            LedgerImage(row_map[image.row_index], image.column_name, image.data, image.width, image.height)
            for image in data.images
            if image.row_index in row_map
        ]
        return LedgerData(
            data.headers[:],
            rows,
            data.source_path,
            data.date_label,
            data.metadata.copy(),
            images,
        )

    def _renumber_rows(self, headers: list[str], rows: list[list[object]]) -> None:
        target = next((name for name in ("编号", "序号") if name in headers), None)
        if not target:
            return
        number_idx = headers.index(target)
        for row_number, row in enumerate(rows, start=1):
            if number_idx < len(row):
                row[number_idx] = row_number
