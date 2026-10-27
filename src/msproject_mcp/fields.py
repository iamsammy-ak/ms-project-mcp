"""Field registry: one place that maps stable field names to COM properties.

Tools accept a ``fields`` list instead of exposing a bespoke parameter for every
Project property. That keeps the tool surface small while still covering the
whole object model, including Text1-30 / Number1-20 / Flag1-20 custom fields.
"""

from __future__ import annotations

import difflib
from dataclasses import dataclass
from typing import Any

from .errors import InvalidInput
from .util import parse_datetime, parse_duration, to_bool

# PjConstraintType
CONSTRAINT_TYPES = {
    "ASAP": 0,
    "ALAP": 1,
    "MSO": 2,
    "MFO": 3,
    "SNET": 4,
    "SNLT": 5,
    "FNET": 6,
    "FNLT": 7,
}

TASK_TYPES = {"fixed_units": 0, "fixed_duration": 1, "fixed_work": 2}
RESOURCE_TYPES = {"work": 0, "material": 1, "cost": 2}


@dataclass(frozen=True)
class Field:
    name: str
    attr: str
    kind: str = "raw"
    writable: bool = True


def _custom(prefix: str, attr: str, count: int, kind: str) -> dict[str, Field]:
    return {
        f"{prefix}{index}": Field(f"{prefix}{index}", f"{attr}{index}", kind)
        for index in range(1, count + 1)
    }


TASK_FIELDS: dict[str, Field] = {
    "name": Field("name", "Name", "str"),
    "notes": Field("notes", "Notes", "str"),
    "duration": Field("duration", "Duration", "duration"),
    "start": Field("start", "Start", "date"),
    "finish": Field("finish", "Finish", "date"),
    "percent_complete": Field("percent_complete", "PercentComplete", "percent"),
    "percent_work_complete": Field("percent_work_complete", "PercentWorkComplete", "percent"),
    "actual_start": Field("actual_start", "ActualStart", "date"),
    "actual_finish": Field("actual_finish", "ActualFinish", "date"),
    "remaining_duration": Field("remaining_duration", "RemainingDuration", "duration"),
    "work": Field("work", "Work", "duration"),
    "actual_work": Field("actual_work", "ActualWork", "duration"),
    "remaining_work": Field("remaining_work", "RemainingWork", "duration"),
    "overtime_work": Field("overtime_work", "OvertimeWork", "duration"),
    "leveling_delay": Field("leveling_delay", "LevelingDelay", "duration"),
    "cost": Field("cost", "Cost", "float", writable=False),
    "fixed_cost": Field("fixed_cost", "FixedCost", "float"),
    "priority": Field("priority", "Priority", "int"),
    "milestone": Field("milestone", "Milestone", "bool"),
    "summary": Field("summary", "Summary", "bool", writable=False),
    "outline_level": Field("outline_level", "OutlineLevel", "int", writable=False),
    "wbs": Field("wbs", "WBS", "str", writable=False),
    "critical": Field("critical", "Critical", "bool", writable=False),
    "total_slack": Field("total_slack", "TotalSlack", "duration", writable=False),
    "free_slack": Field("free_slack", "FreeSlack", "duration", writable=False),
    "early_start": Field("early_start", "EarlyStart", "date", writable=False),
    "early_finish": Field("early_finish", "EarlyFinish", "date", writable=False),
    "late_start": Field("late_start", "LateStart", "date", writable=False),
    "late_finish": Field("late_finish", "LateFinish", "date", writable=False),
    "constraint_type": Field("constraint_type", "ConstraintType", "constraint"),
    "constraint_date": Field("constraint_date", "ConstraintDate", "date"),
    "deadline": Field("deadline", "Deadline", "date"),
    "manual": Field("manual", "Manual", "bool"),
    "task_type": Field("task_type", "Type", "task_type"),
    "effort_driven": Field("effort_driven", "EffortDriven", "bool"),
    "ignore_resource_calendar": Field(
        "ignore_resource_calendar", "IgnoreResourceCalendar", "bool"
    ),
    "confirmed": Field("confirmed", "Confirmed", "bool"),
    "active": Field("active", "Active", "bool"),
    "calendar": Field("calendar", "Calendar", "calendar"),
    "hyperlink": Field("hyperlink", "Hyperlink", "str"),
    "hyperlink_text": Field("hyperlink_text", "HyperlinkSubAddress", "str"),
    "predecessors": Field("predecessors", "Predecessors", "str"),
    "successors": Field("successors", "Successors", "str", writable=False),
    "resource_names": Field("resource_names", "ResourceNames", "str"),
    "baseline_start": Field("baseline_start", "BaselineStart", "date", writable=False),
    "baseline_finish": Field("baseline_finish", "BaselineFinish", "date", writable=False),
    "baseline_duration": Field(
        "baseline_duration", "BaselineDuration", "duration", writable=False
    ),
    "baseline_work": Field("baseline_work", "BaselineWork", "duration", writable=False),
    "baseline_cost": Field("baseline_cost", "BaselineCost", "float", writable=False),
    "bcws": Field("bcws", "BCWS", "float", writable=False),
    "bcwp": Field("bcwp", "BCWP", "float", writable=False),
    "acwp": Field("acwp", "ACWP", "float", writable=False),
    "sv": Field("sv", "SV", "float", writable=False),
    "cv": Field("cv", "CV", "float", writable=False),
    "spi": Field("spi", "SPI", "float", writable=False),
    "cpi": Field("cpi", "CPI", "float", writable=False),
}
TASK_FIELDS.update(_custom("text", "Text", 30, "str"))
TASK_FIELDS.update(_custom("number", "Number", 20, "float"))
TASK_FIELDS.update(_custom("flag", "Flag", 20, "bool"))
TASK_FIELDS.update(_custom("date", "Date", 10, "date"))
TASK_FIELDS.update(_custom("cost", "Cost", 10, "float"))
TASK_FIELDS.update(_custom("duration", "Duration", 10, "duration"))


PROJECT_FIELDS: dict[str, Field] = {
    "title": Field("title", "Title", "str"),
    "subject": Field("subject", "Subject", "str"),
    "author": Field("author", "Author", "str"),
    "company": Field("company", "Company", "str"),
    "manager": Field("manager", "Manager", "str"),
    "category": Field("category", "Category", "str"),
    "keywords": Field("keywords", "Keywords", "str"),
    "comments": Field("comments", "Comments", "str"),
    "start": Field("start", "ProjectStart", "date"),
    "computed_start": Field("computed_start", "Start", "date", writable=False),
    "finish": Field("finish", "ProjectFinish", "date", writable=False),
    "status_date": Field("status_date", "StatusDate", "date"),
    "current_date": Field("current_date", "CurrentDate", "date"),
    "calendar": Field("calendar", "Calendar", "calendar"),
    "schedule_from_start": Field("schedule_from_start", "ScheduleFromStart", "bool"),
    "minutes_per_day": Field("minutes_per_day", "MinutesPerDay", "int"),
    "minutes_per_week": Field("minutes_per_week", "MinutesPerWeek", "int"),
    "default_start_time": Field("default_start_time", "DefaultStartTime", "date"),
    "default_finish_time": Field("default_finish_time", "DefaultFinishTime", "date"),
    "currency_code": Field("currency_code", "CurrencyCode", "str"),
    "currency_symbol": Field("currency_symbol", "CurrencySymbol", "str"),
    "priority": Field("priority", "Priority", "int"),
    "hyperlink": Field("hyperlink", "Hyperlink", "str"),
}


RESOURCE_FIELDS: dict[str, Field] = {
    "name": Field("name", "Name", "str"),
    "type": Field("type", "Type", "resource_type"),
    "max_units": Field("max_units", "MaxUnits", "float"),
    "standard_rate": Field("standard_rate", "StandardRate", "float"),
    "overtime_rate": Field("overtime_rate", "OvertimeRate", "float"),
    "cost_per_use": Field("cost_per_use", "CostPerUse", "float"),
    "group": Field("group", "Group", "str"),
    "email_address": Field("email_address", "EmailAddress", "str"),
    "initials": Field("initials", "Initials", "str"),
    "material_label": Field("material_label", "MaterialLabel", "str"),
    "calendar": Field("calendar", "Calendar", "calendar"),
    "notes": Field("notes", "Notes", "str"),
    "code": Field("code", "Code", "str"),
    "budget": Field("budget", "Budget", "bool"),
    "cost": Field("cost", "Cost", "float", writable=False),
    "work": Field("work", "Work", "duration", writable=False),
    "actual_work": Field("actual_work", "ActualWork", "duration", writable=False),
    "remaining_work": Field("remaining_work", "RemainingWork", "duration", writable=False),
    "percent_work_complete": Field(
        "percent_work_complete", "PercentWorkComplete", "percent", writable=False
    ),
}
RESOURCE_FIELDS.update(_custom("text", "Text", 30, "str"))
RESOURCE_FIELDS.update(_custom("number", "Number", 20, "float"))
RESOURCE_FIELDS.update(_custom("flag", "Flag", 20, "bool"))
RESOURCE_FIELDS.update(_custom("date", "Date", 10, "date"))
RESOURCE_FIELDS.update(_custom("cost", "Cost", 10, "float"))


DEFAULT_TASK_FIELDS = [
    "unique_id",
    "id",
    "name",
    "outline_level",
    "summary",
    "milestone",
    "duration",
    "start",
    "finish",
    "percent_complete",
    "critical",
    "predecessors",
    "resource_names",
]

DEFAULT_RESOURCE_FIELDS = [
    "unique_id",
    "id",
    "name",
    "type",
    "max_units",
    "standard_rate",
    "group",
    "calendar",
]

_ALIASES = {
    "task_name": "name",
    "taskname": "name",
    "complete": "percent_complete",
    "pct_complete": "percent_complete",
    "percent_completed": "percent_complete",
    "predecessor": "predecessors",
    "successor": "successors",
    "resources": "resource_names",
    "start_date": "start",
    "end_date": "finish",
    "finish_date": "finish",
    "constraint": "constraint_type",
    "resource_name": "name",
    "notes_text": "notes",
}


def normalize_name(name: Any) -> str:
    text = str(name).strip().lower()
    for prefix in ("task.", "project.", "resource."):
        if text.startswith(prefix):
            text = text[len(prefix) :]
    text = text.replace("-", "_").replace(" ", "_")
    return _ALIASES.get(text, text)


def lookup(registry: dict[str, Field], name: Any) -> Field:
    key = normalize_name(name)
    field = registry.get(key)
    if field is not None:
        return field
    candidates = difflib.get_close_matches(key, registry.keys(), n=3, cutoff=0.6)
    hint = f" Did you mean: {', '.join(candidates)}?" if candidates else ""
    raise InvalidInput(f"Unknown field {name!r}.{hint}")


def normalize_field_list(
    registry: dict[str, Field], names: Any, default: list[str]
) -> list[str]:
    if names is None:
        return list(default)
    if isinstance(names, str):
        names = [names]
    result: list[str] = []
    for name in names:
        key = normalize_name(name)
        if key in {"unique_id", "id"}:
            if key not in result:
                result.append(key)
            continue
        field = lookup(registry, key)
        if field.name not in result:
            result.append(field.name)
    return result


def coerce(field: Field, value: Any, minutes_per_day: float, minutes_per_week: float) -> Any:
    if value is None:
        return None
    kind = field.kind
    if kind == "str":
        return str(value)
    if kind == "int":
        return int(value)
    if kind == "float":
        return float(value)
    if kind == "bool":
        return to_bool(value)
    if kind == "date":
        return parse_datetime(value)
    if kind == "duration":
        return parse_duration(value, minutes_per_day, minutes_per_week)
    if kind == "percent":
        number = float(value)
        if 0.0 < number < 1.0 and not float(number).is_integer():
            number *= 100.0
        if not 0.0 <= number <= 100.0:
            raise InvalidInput(f"{field.name} must be between 0 and 100 (got {value!r}).")
        return int(round(number))
    if kind == "constraint":
        return _enum(CONSTRAINT_TYPES, value, "constraint type")
    if kind == "task_type":
        return _enum(TASK_TYPES, value, "task type")
    if kind == "resource_type":
        return _enum(RESOURCE_TYPES, value, "resource type")
    return value


def _enum(table: dict[str, int], value: Any, label: str) -> int:
    if isinstance(value, int) and not isinstance(value, bool):
        if value in table.values():
            return value
        raise InvalidInput(f"{value} is not a valid {label}. Use one of {list(table)}.")
    key = str(value).strip().lower().replace(" ", "_").replace("-", "_")
    if key in table:
        return table[key]
    for name, number in table.items():
        if name.lower() == key:
            return number
    raise InvalidInput(f"{value!r} is not a valid {label}. Use one of {list(table)}.")


def enum_name(table: dict[str, int], value: Any) -> str | None:
    try:
        number = int(value)
    except (TypeError, ValueError):
        return None
    for name, candidate in table.items():
        if candidate == number:
            return name
    return None


def catalog(registry: dict[str, Field], writable_only: bool = False) -> dict[str, list[str]]:
    writable = sorted(name for name, field in registry.items() if field.writable)
    read_only = sorted(name for name, field in registry.items() if not field.writable)
    result = {"writable": writable, "read_only": read_only}
    if writable_only:
        result.pop("read_only")
    return result
