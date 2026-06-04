"""Normalized parser models."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, time


@dataclass
class ScheduleRecord:
    employee_display_name: str
    work_date: date
    shift_kind: str
    start_time: time
    end_time: time
    job_code: str
    role_group: str
    is_clinical_leader: bool = False
    has_lead_pay: bool = False
    source_evidence: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    source_sheet: str | None = None
    source_row: int | None = None


@dataclass
class ParsedSchedule:
    records: list[ScheduleRecord]
    warnings: list[str] = field(default_factory=list)

    @property
    def date_range(self) -> tuple[date, date] | None:
        if not self.records:
            return None
        dates = [record.work_date for record in self.records]
        return min(dates), max(dates)

