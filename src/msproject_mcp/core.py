"""The Microsoft Project domain layer.

Every public method here is safe to call from any thread: it packages the COM
work into a closure that :class:`~msproject_mcp.bridge.ComBridge` runs on the
single STA thread.

Task and resource identifiers: tools accept either ``UniqueID`` (stable across
edits) or the visible ``ID``/row number. UniqueID is checked first.
"""

from __future__ import annotations

import csv
import datetime as dt
import json
import os
import re
from contextlib import contextmanager
from typing import Any, Iterable

import win32com.client

from . import config
from .bridge import ComBridge
from .errors import (
    FieldNotWritable,
    InvalidInput,
    MSProjectMCPError,
    NoProjectOpen,
    NotFound,
    ProjectNotRunning,
)
from .fields import (
    CONSTRAINT_TYPES,
    DEFAULT_RESOURCE_FIELDS,
    DEFAULT_TASK_FIELDS,
    PROJECT_FIELDS,
    RESOURCE_FIELDS,
    TASK_FIELDS,
    coerce,
    enum_name,
    lookup,
    normalize_field_list,
    normalize_name,
)
from .util import (
    format_datetime,
    format_duration,
    parse_datetime,
    parse_duration,
)

_LINK_RE = re.compile(
    r"(\d+)\s*(FS|SS|FF|SF)?\s*([+-]?\s*\d+(?:\.\d+)?\s*[a-zA-Z]+)?", re.IGNORECASE
)
_SAVE_FORMATS = {"mpp": 0, "mpt": 11}
_WEEKDAYS = {
    "sunday": 1,
    "monday": 2,
    "tuesday": 3,
    "wednesday": 4,
    "thursday": 5,
    "friday": 6,
    "saturday": 7,
}


def _safe(getter, default=None):
    try:
        return getter()
    except Exception:
        return default


def _as_float(value):
    """Coerce a COM value to float.

    Rate and currency properties are Variants: Project returns them formatted for
    the current locale, e.g. "₹ 85.00/hr", while accepting a plain number on
    write. Anything unparseable comes back as None so the caller can fall back to
    the raw string.
    """
    if value is None:
        return None
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text or text.upper() == "NA":
        return None
    cleaned = re.sub(r"[^0-9eE+\-.]", "", text)
    try:
        return float(cleaned)
    except ValueError:
        return None


@contextmanager
def _quiet(app):
    """Suppress alerts and screen updates, then restore the user's settings."""
    alerts = _safe(lambda: app.DisplayAlerts)
    updating = _safe(lambda: app.ScreenUpdating)
    _safe(lambda: setattr(app, "DisplayAlerts", False))
    _safe(lambda: setattr(app, "ScreenUpdating", False))
    try:
        yield
    finally:
        if updating is not None:
            _safe(lambda: setattr(app, "ScreenUpdating", updating))
        if alerts is not None:
            _safe(lambda: setattr(app, "DisplayAlerts", alerts))


def _format_lag(lag: Any, minutes_per_day: float) -> str:
    if lag is None or lag == "":
        return ""
    if isinstance(lag, str):
        text = lag.strip()
        if not text:
            return ""
        return text if text[0] in "+-" else f"+{text}"
    minutes = parse_duration(lag, minutes_per_day)
    if minutes == 0:
        return ""
    text = format_duration(abs(minutes), minutes_per_day)
    return f"+{text}" if minutes > 0 else f"-{text}"


def _parse_links(text: Any) -> list[dict[str, Any]]:
    links: list[dict[str, Any]] = []
    for raw in str(text or "").split(","):
        raw = raw.strip()
        if not raw:
            continue
        match = _LINK_RE.search(raw)
        if not match:
            continue
        links.append(
            {
                "task_id": int(match.group(1)),
                "type": (match.group(2) or "FS").upper(),
                "lag": (match.group(3) or "").replace(" ", "") or "0d",
            }
        )
    return links


class MSProjectClient:
    def __init__(self) -> None:
        self._bridge = ComBridge()
        self._app = None
        self._proj = None
        self._attached = False

    # ------------------------------------------------------------------
    # Connection and project lifecycle
    # ------------------------------------------------------------------

    def _run(self, fn, timeout: float | None = None):
        return self._bridge.call(fn, timeout=timeout)

    def _app_ref(self, launch: bool = True):
        if self._app is not None:
            try:
                self._app.Projects.Count  # noqa: B018 - liveness probe
                return self._app
            except Exception:
                self._app = None
                self._proj = None

        try:
            self._app = win32com.client.GetActiveObject("MSProject.Application")
            self._attached = True
        except Exception:
            if not launch:
                raise ProjectNotRunning(
                    "Microsoft Project is not running. Ask me to open or create a "
                    "project and I will start it."
                ) from None
            self._app = win32com.client.Dispatch("MSProject.Application")
            self._attached = False
            if config.VISIBLE:
                _safe(lambda: setattr(self._app, "Visible", True))
        return self._app

    def _find_project(self, app, selector: Any):
        count = _safe(lambda: int(app.Projects.Count), 0)
        if count == 0:
            raise NoProjectOpen(
                "No project file is open in Microsoft Project. Use create_project "
                "or open_project first."
            )

        text = str(selector).strip()
        if text.isdigit():
            index = int(text)
            if 1 <= index <= count:
                return app.Projects(index)
            raise NotFound(f"Project index {index} is out of range (1-{count}).")

        needle = text.replace("/", "\\").lower()
        matches = []
        for index in range(1, count + 1):
            project = app.Projects(index)
            name = str(_safe(lambda: project.Name, "") or "")
            full = str(_safe(lambda: project.FullName, "") or "")
            if needle in name.lower() or needle in full.lower():
                matches.append(project)

        if not matches:
            names = [
                str(_safe(lambda: app.Projects(i).Name, f"project {i}"))
                for i in range(1, count + 1)
            ]
            raise NotFound(f"No open project matches {selector!r}. Open: {names}")

        for project in matches:
            if str(_safe(lambda: project.Name, "")).lower() == needle:
                return project
        return matches[0]

    def _project_ref(self, project: Any = None):
        # Read tools must not silently start Project: only create_project and
        # open_project are allowed to launch it.
        app = self._app_ref(launch=False)
        if project not in (None, ""):
            proj = self._find_project(app, project)
            _safe(lambda: proj.Activate())
            self._proj = proj
            return proj

        if self._proj is not None:
            try:
                self._proj.Name  # noqa: B018 - liveness probe
                return self._proj
            except Exception:
                self._proj = None

        if _safe(lambda: int(app.Projects.Count), 0) == 0:
            raise NoProjectOpen(
                "No project file is open in Microsoft Project. Use create_project "
                "or open_project first."
            )
        self._proj = app.ActiveProject
        return self._proj

    @staticmethod
    def _activate(app, proj) -> None:
        """Make ``proj`` the active project without stealing focus unnecessarily.

        Application-level methods (FileSave, FileSaveAs, LevelNow, BaselineSave,
        DocumentExport, CalculateProject) always act on the active project, so
        they must be aimed at the project the caller asked for.
        """
        active = _safe(lambda: str(app.ActiveProject.Name), "")
        if active != _safe(lambda: str(proj.Name), ""):
            _safe(lambda: proj.Activate())

    @staticmethod
    def _minutes_per_day(proj) -> float:
        return float(_safe(lambda: proj.MinutesPerDay, 480) or 480)

    @staticmethod
    def _minutes_per_week(proj) -> float:
        return float(_safe(lambda: proj.MinutesPerWeek, 2400) or 2400)

    def _project_summary(self, proj) -> dict[str, Any]:
        return {
            "name": _safe(lambda: str(proj.Name)),
            "path": _safe(lambda: str(proj.FullName)),
            "title": _safe(lambda: str(proj.Title)),
            "tasks": _safe(lambda: int(proj.Tasks.Count), 0),
            "resources": _safe(lambda: int(proj.Resources.Count), 0),
            "start": format_datetime(_safe(lambda: proj.ProjectStart)),
            "finish": format_datetime(_safe(lambda: proj.ProjectFinish)),
            "status_date": format_datetime(_safe(lambda: proj.StatusDate)),
            "calendar": _safe(lambda: str(proj.Calendar.Name)),
            "minutes_per_day": _safe(lambda: int(proj.MinutesPerDay), 480),
            "schedule_from_start": _safe(lambda: bool(proj.ScheduleFromStart)),
        }

    def status(self) -> dict[str, Any]:
        def job():
            info: dict[str, Any] = {
                "running": False,
                "vba_enabled": config.ALLOW_VBA,
                "max_rows_per_call": config.MAX_ROWS,
            }
            try:
                app = self._app_ref(launch=False)
            except ProjectNotRunning:
                return info

            info["running"] = True
            info["visible"] = _safe(lambda: bool(app.Visible))
            info["version"] = _safe(lambda: str(app.Version))
            info["launched_by_server"] = not self._attached

            count = _safe(lambda: int(app.Projects.Count), 0)
            info["open_project_count"] = count
            projects = []
            for index in range(1, count + 1):
                project = app.Projects(index)
                projects.append(
                    {
                        "index": index,
                        "name": _safe(lambda: str(project.Name)),
                        "path": _safe(lambda: str(project.FullName)),
                        "tasks": _safe(lambda: int(project.Tasks.Count), 0),
                    }
                )
            info["projects"] = projects
            info["active_project"] = _safe(lambda: str(app.ActiveProject.Name))
            return info

        return self._run(job)

    def list_projects(self) -> dict[str, Any]:
        def job():
            app = self._app_ref(launch=False)
            count = int(app.Projects.Count)
            active = _safe(lambda: str(app.ActiveProject.Name), "")
            projects = []
            for index in range(1, count + 1):
                project = app.Projects(index)
                name = _safe(lambda: str(project.Name), "")
                projects.append(
                    {
                        "index": index,
                        "name": name,
                        "path": _safe(lambda: str(project.FullName)),
                        "tasks": _safe(lambda: int(project.Tasks.Count), 0),
                        "resources": _safe(lambda: int(project.Resources.Count), 0),
                        "start": format_datetime(_safe(lambda: project.ProjectStart)),
                        "finish": format_datetime(_safe(lambda: project.ProjectFinish)),
                        "active": name == active,
                    }
                )
            return {"count": count, "active": active, "projects": projects}

        return self._run(job)

    def use_project(self, selector: Any) -> dict[str, Any]:
        def job():
            app = self._app_ref(launch=False)
            project = self._find_project(app, selector)
            with _quiet(app):
                project.Activate()
            self._proj = project
            summary = self._project_summary(project)
            summary["active"] = True
            return summary

        return self._run(job)

    def open_project(self, path: str, read_only: bool = False) -> dict[str, Any]:
        target = os.path.abspath(os.path.expanduser(str(path)))
        if not os.path.exists(target):
            raise NotFound(f"File not found: {target}")

        def job():
            app = self._app_ref()
            with _quiet(app):
                try:
                    project = app.FileOpen(Name=target, ReadOnly=bool(read_only))
                except TypeError:
                    project = app.FileOpen(target)
            if project is None:
                project = app.ActiveProject
            self._proj = project
            summary = self._project_summary(project)
            summary["read_only"] = bool(read_only)
            return summary

        return self._run(job, timeout=300)

    def create_project(self, title: str | None = None, start: str | None = None) -> dict[str, Any]:
        def job():
            app = self._app_ref()
            with _quiet(app):
                app.FileNew()
                project = app.ActiveProject
                if title:
                    project.Title = str(title)
                if start:
                    value = parse_datetime(start)
                    try:
                        project.ProjectStart = value
                    except Exception:
                        project.Start = value
            self._proj = project
            return self._project_summary(project)

        return self._run(job, timeout=180)

    def save_project(self) -> dict[str, Any]:
        def job():
            app = self._app_ref()
            project = self._project_ref()
            with _quiet(app):
                self._activate(app, project)
                app.FileSave()
            return self._project_summary(project)

        return self._run(job, timeout=300)

    def save_project_as(self, path: str, fmt: str = "mpp") -> dict[str, Any]:
        fmt_key = str(fmt or "mpp").lower().strip()
        if fmt_key not in _SAVE_FORMATS:
            raise InvalidInput(
                f"Unsupported format {fmt!r}. Use one of {sorted(_SAVE_FORMATS)}."
            )
        target = os.path.abspath(os.path.expanduser(str(path)))

        def job():
            app = self._app_ref()
            project = self._project_ref()
            with _quiet(app):
                self._activate(app, project)
                app.FileSaveAs(
                    Name=target,
                    Format=_SAVE_FORMATS[fmt_key],
                    Backup=False,
                    ReadOnly=False,
                )
            self._proj = app.ActiveProject
            summary = self._project_summary(self._proj)
            summary["saved_to"] = target
            if fmt_key == "mpt":
                summary["warning"] = (
                    "Saved as a template. The open project now points at the .mpt, "
                    "exactly as Project's own Save As does."
                )
            return summary

        return self._run(job, timeout=600)

    def close_project(self, save: bool = False, project: Any = None) -> dict[str, Any]:
        def job():
            app = self._app_ref()
            proj = self._project_ref(project)
            with _quiet(app):
                self._activate(app, proj)
            name = _safe(lambda: str(proj.Name), "")
            with _quiet(app):
                app.FileClose(Save=1 if save else 0)
            self._proj = None
            return {"closed": name, "saved": bool(save)}

        return self._run(job, timeout=300)

    def quit_application(self, save: bool = False) -> dict[str, Any]:
        def job():
            app = self._app_ref(launch=False)
            with _quiet(app):
                if save:
                    _safe(lambda: app.FileSave())
                app.Quit()
            self._app = None
            self._proj = None
            return {"quit": True, "saved": bool(save)}

        return self._run(job, timeout=300)

    def set_visible(self, visible: bool) -> dict[str, Any]:
        def job():
            app = self._app_ref(launch=False)
            app.Visible = bool(visible)
            return {"visible": bool(app.Visible)}

        return self._run(job)

    # ------------------------------------------------------------------
    # Field plumbing
    # ------------------------------------------------------------------

    @staticmethod
    def _convert_value(field, value, mpd: float, mpw: float) -> dict[str, Any]:
        if value is None:
            return {field.name: None}

        kind = field.kind
        if kind == "duration":
            try:
                minutes = float(value)
            except (TypeError, ValueError):
                try:
                    minutes = parse_duration(value, mpd, mpw)
                except (TypeError, ValueError):
                    return {field.name: str(value)}
            if minutes is None:
                return {field.name: None}
            return {
                field.name: format_duration(minutes, mpd, mpw),
                f"{field.name}_minutes": round(minutes, 2),
            }
        if kind == "date":
            return {field.name: format_datetime(value)}
        if kind == "bool":
            return {field.name: bool(value)}
        if kind == "int":
            number = _as_float(value)
            return {field.name: int(number) if number is not None else None}
        if kind == "float":
            number = _as_float(value)
            result = {field.name: round(number, 4) if number is not None else None}
            if number is not None and isinstance(value, str):
                result[f"{field.name}_text"] = str(value)
            return result
        if kind == "percent":
            number = _as_float(value)
            return {field.name: int(round(number)) if number is not None else None}
        if kind == "constraint":
            return {field.name: enum_name(CONSTRAINT_TYPES, value)}
        if kind == "task_type":
            from .fields import TASK_TYPES

            return {field.name: enum_name(TASK_TYPES, value)}
        if kind == "resource_type":
            from .fields import RESOURCE_TYPES

            return {field.name: enum_name(RESOURCE_TYPES, value)}
        if kind == "calendar":
            return {field.name: _safe(lambda: str(value.Name), str(value))}
        return {field.name: str(value)}

    def _read_fields(
        self,
        obj,
        registry: dict,
        names: Iterable[str],
        mpd: float,
        mpw: float,
    ) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for name in names:
            if name == "unique_id":
                result["unique_id"] = _safe(lambda: int(obj.UniqueID))
                continue
            if name == "id":
                result["id"] = _safe(lambda: int(obj.ID))
                continue
            field = registry[name]
            value = _safe(lambda: getattr(obj, field.attr))
            result.update(self._convert_value(field, value, mpd, mpw))
        return result

    def _write_fields(
        self,
        obj,
        registry: dict,
        values: dict[str, Any],
        mpd: float,
        mpw: float,
        proj=None,
    ) -> list[str]:
        applied: list[str] = []
        for raw_name, raw_value in values.items():
            name = normalize_name(raw_name)
            if name in {"unique_id", "id"}:
                raise InvalidInput(f"{name} is managed by Project and cannot be written.")
            if name.endswith("_minutes"):
                base = name[: -len("_minutes")]
                if base in registry and registry[base].kind == "duration":
                    name = base
            field = lookup(registry, name)
            if not field.writable:
                raise FieldNotWritable(
                    f"{field.name} is read-only in Microsoft Project. "
                    "Change the inputs that drive it instead (for example duration, "
                    "links, or resource assignments)."
                )
            if field.kind == "calendar" and proj is not None:
                coerced = self._resolve_calendar(proj, raw_value)
            else:
                coerced = coerce(field, raw_value, mpd, mpw)
            setattr(obj, field.attr, coerced)
            applied.append(field.name)
        return applied

    @staticmethod
    def _resolve_calendar(proj, value):
        if value is None:
            return None
        needle = str(value).strip().lower()
        for collection_name in ("Calendars", "BaseCalendars"):
            collection = _safe(lambda: getattr(proj, collection_name))
            if collection is None:
                continue
            count = _safe(lambda: int(collection.Count), 0)
            for index in range(1, count + 1):
                calendar = _safe(lambda: collection(index))
                if calendar is not None and str(_safe(lambda: calendar.Name, "")).lower() == needle:
                    return calendar
        raise NotFound(f"Calendar {value!r} was not found in this project.")

    def _calendar(self, proj, name: str):
        needle = str(name).strip().lower()
        collection = _safe(lambda: proj.BaseCalendars)
        count = _safe(lambda: int(collection.Count), 0) if collection is not None else 0
        for index in range(1, count + 1):
            calendar = collection(index)
            if str(_safe(lambda: calendar.Name, "")).lower() == needle:
                return calendar
        names = [
            str(_safe(lambda: collection(i).Name, "")) for i in range(1, count + 1)
        ]
        raise NotFound(f"Calendar {name!r} not found. Available: {names}")

    @staticmethod
    def _set_outline_level(app, task, level: int) -> int:
        """Move a task to an outline level.

        ``Task.OutlineLevel`` is writable on modern Project builds; older ones
        only support the OutlineIndent/OutlineOutdent commands, which act on the
        selected row.
        """
        target = max(1, min(10, int(level)))
        try:
            task.OutlineLevel = target
            if int(_safe(lambda: task.OutlineLevel, target)) == target:
                return target
        except Exception:
            pass

        _safe(lambda: app.SelectRow(task.ID, False))
        current = int(_safe(lambda: task.OutlineLevel, 1))
        for _ in range(max(0, target - current)):
            try:
                app.OutlineIndent()
            except Exception:
                break
        for _ in range(max(0, current - target)):
            try:
                app.OutlineOutdent()
            except Exception:
                break
        return int(_safe(lambda: task.OutlineLevel, current))

    @staticmethod
    def _find_task(proj, task_id: Any):
        try:
            wanted = int(task_id)
        except (TypeError, ValueError):
            raise InvalidInput(f"Task id must be an integer, got {task_id!r}.") from None
        by_id = None
        for task in proj.Tasks:
            if task is None:
                continue
            unique = _safe(lambda: int(task.UniqueID))
            if unique == wanted:
                return task
            if by_id is None and _safe(lambda: int(task.ID)) == wanted:
                by_id = task
        return by_id

    @staticmethod
    def _find_resource(proj, selector: Any):
        text = str(selector).strip()
        wanted = int(text) if text.isdigit() else None
        by_id = None
        for resource in proj.Resources:
            if resource is None:
                continue
            if wanted is not None:
                if _safe(lambda: int(resource.UniqueID)) == wanted:
                    return resource
                if by_id is None and _safe(lambda: int(resource.ID)) == wanted:
                    by_id = resource
            elif str(_safe(lambda: resource.Name, "")).strip().lower() == text.lower():
                return resource
        return by_id

    # ------------------------------------------------------------------
    # Project properties
    # ------------------------------------------------------------------

    def get_project(self, fields: Any = None, project: Any = None) -> dict[str, Any]:
        def job():
            proj = self._project_ref(project)
            names = normalize_field_list(
                PROJECT_FIELDS, fields, [f.name for f in PROJECT_FIELDS.values()]
            )
            mpd, mpw = self._minutes_per_day(proj), self._minutes_per_week(proj)
            summary = self._project_summary(proj)
            summary["fields"] = self._read_fields(proj, PROJECT_FIELDS, names, mpd, mpw)
            return summary

        return self._run(job)

    def update_project(self, values: dict[str, Any], project: Any = None) -> dict[str, Any]:
        if not isinstance(values, dict) or not values:
            raise InvalidInput("Provide at least one project field to update.")
        # Guard against the most common field mix-up before touching COM.
        for name in values:
            lookup(PROJECT_FIELDS, name)

        def job():
            app = self._app_ref()
            proj = self._project_ref(project)
            mpd, mpw = self._minutes_per_day(proj), self._minutes_per_week(proj)
            names = [
                normalize_name(name)
                for name in values
                if normalize_name(name) in PROJECT_FIELDS
            ]
            before = self._read_fields(proj, PROJECT_FIELDS, names, mpd, mpw)
            with _quiet(app):
                self._write_fields(proj, PROJECT_FIELDS, values, mpd, mpw, proj)
            after = self._read_fields(proj, PROJECT_FIELDS, names, mpd, mpw)
            return {
                "updated": names,
                "project": self._project_summary(proj),
                "changes": {
                    key: {"from": before.get(key), "to": after.get(key)}
                    for key in names
                },
            }

        return self._run(job)

    # ------------------------------------------------------------------
    # Tasks
    # ------------------------------------------------------------------

    def get_tasks(
        self,
        fields: Any = None,
        name_contains: str | None = None,
        outline_level: int | None = None,
        include_summary: bool = True,
        milestones_only: bool = False,
        limit: int | None = None,
        offset: int = 0,
        project: Any = None,
    ) -> dict[str, Any]:
        def job():
            proj = self._project_ref(project)
            mpd, mpw = self._minutes_per_day(proj), self._minutes_per_week(proj)
            names = normalize_field_list(TASK_FIELDS, fields, DEFAULT_TASK_FIELDS)
            cap = min(int(limit or config.MAX_ROWS), config.MAX_ROWS)
            needle = (name_contains or "").strip().lower()

            items: list[dict[str, Any]] = []
            total = 0
            matched = 0
            for task in proj.Tasks:
                if task is None:
                    continue
                total += 1
                if not include_summary and _safe(lambda: bool(task.Summary), False):
                    continue
                if outline_level is not None and _safe(
                    lambda: int(task.OutlineLevel), 0
                ) != int(outline_level):
                    continue
                if milestones_only and not _safe(lambda: bool(task.Milestone), False):
                    continue
                if needle and needle not in str(_safe(lambda: task.Name, "")).lower():
                    continue
                matched += 1
                if matched <= int(offset):
                    continue
                if len(items) >= cap:
                    continue
                items.append(self._read_fields(task, TASK_FIELDS, names, mpd, mpw))

            return {
                "project": _safe(lambda: str(proj.Name)),
                "total_tasks": total,
                "matched": matched,
                "returned": len(items),
                "offset": int(offset),
                "limit": cap,
                "truncated": matched - int(offset) > len(items),
                "items": items,
            }

        return self._run(job, timeout=300)

    def get_task(self, unique_id: Any, fields: Any = None, project: Any = None) -> dict[str, Any]:
        def job():
            proj = self._project_ref(project)
            task = self._find_task(proj, unique_id)
            if task is None:
                raise NotFound(f"No task with id {unique_id} in {_safe(lambda: proj.Name)}.")
            mpd, mpw = self._minutes_per_day(proj), self._minutes_per_week(proj)
            names = normalize_field_list(TASK_FIELDS, fields, DEFAULT_TASK_FIELDS)
            data = self._read_fields(task, TASK_FIELDS, names, mpd, mpw)
            data["project"] = _safe(lambda: str(proj.Name))
            return data

        return self._run(job)

    def create_tasks(self, tasks: list[dict[str, Any]], project: Any = None) -> dict[str, Any]:
        if not isinstance(tasks, list) or not tasks:
            raise InvalidInput("Provide a non-empty list of task objects.")
        for item in tasks:
            if not isinstance(item, dict):
                raise InvalidInput("Each task must be a JSON object.")
            if not str(item.get("name") or "").strip():
                raise InvalidInput("Each task needs a non-empty 'name'.")

        def job():
            app = self._app_ref()
            proj = self._project_ref(project)
            mpd, mpw = self._minutes_per_day(proj), self._minutes_per_week(proj)
            created = []
            with _quiet(app):
                for item in tasks:
                    name = str(item["name"]).strip()
                    task = proj.Tasks.Add(name)
                    values = dict(item)
                    values.pop("name", None)
                    indent = int(values.pop("indent", 0) or 0)
                    # Tasks.Add creates *manually* scheduled tasks in Project 16,
                    # which freeze their dates and ignore links. Automation almost
                    # always wants a calculated schedule, so default to automatic
                    # unless the caller explicitly asked for manual.
                    if not any(normalize_name(key) == "manual" for key in values):
                        values = {"manual": False, **values}
                    self._write_fields(task, TASK_FIELDS, values, mpd, mpw, proj)
                    self._set_outline_level(app, task, 1 + indent)
                    created.append(task)
                _safe(lambda: app.CalculateProject())

            names = normalize_field_list(TASK_FIELDS, None, DEFAULT_TASK_FIELDS)
            return {
                "created": len(created),
                "items": [
                    self._read_fields(task, TASK_FIELDS, names, mpd, mpw)
                    for task in created
                ],
            }

        return self._run(job, timeout=600)

    def update_tasks(
        self,
        updates: list[dict[str, Any]],
        dry_run: bool = False,
        project: Any = None,
    ) -> dict[str, Any]:
        if not isinstance(updates, list) or not updates:
            raise InvalidInput("Provide a non-empty list of updates.")

        def job():
            app = self._app_ref()
            proj = self._project_ref(project)
            mpd, mpw = self._minutes_per_day(proj), self._minutes_per_week(proj)
            results = []

            with _quiet(app):
                for update in updates:
                    if not isinstance(update, dict):
                        raise InvalidInput("Each update must be a JSON object.")
                    task_id = update.get("unique_id", update.get("id", update.get("task")))
                    if task_id is None:
                        raise InvalidInput(
                            "Each update needs 'unique_id' (or 'id')."
                        )
                    values = update.get("set")
                    if values is None:
                        values = {
                            key: value
                            for key, value in update.items()
                            if key not in {"unique_id", "id", "task"}
                        }
                    if not isinstance(values, dict) or not values:
                        raise InvalidInput(
                            f"Update for task {task_id} has no fields to change."
                        )

                    task = self._find_task(proj, task_id)
                    if task is None:
                        raise NotFound(
                            f"No task with id {task_id} in {_safe(lambda: proj.Name)}."
                        )

                    # Validate the field set before writing any value.
                    normalized = {}
                    for raw_name in values:
                        name = normalize_name(raw_name)
                        if name in {"unique_id", "id"}:
                            raise InvalidInput("unique_id/id cannot be written.")
                        if name.endswith("_minutes"):
                            base = name[: -len("_minutes")]
                            if base in TASK_FIELDS and TASK_FIELDS[base].kind == "duration":
                                name = base
                        field = lookup(TASK_FIELDS, name)
                        if not field.writable:
                            raise FieldNotWritable(
                                f"{field.name} is read-only in Microsoft Project."
                            )
                        normalized[field.name] = values[raw_name]

                    before = self._read_fields(
                        task, TASK_FIELDS, list(normalized), mpd, mpw
                    )
                    if dry_run:
                        results.append(
                            {
                                "unique_id": _safe(lambda: int(task.UniqueID)),
                                "name": _safe(lambda: str(task.Name)),
                                "changes": {
                                    name: {"from": before.get(name), "to": value}
                                    for name, value in normalized.items()
                                },
                            }
                        )
                        continue

                    self._write_fields(task, TASK_FIELDS, normalized, mpd, mpw, proj)
                    after = self._read_fields(
                        task, TASK_FIELDS, list(normalized), mpd, mpw
                    )
                    results.append(
                        {
                            "unique_id": _safe(lambda: int(task.UniqueID)),
                            "name": _safe(lambda: str(task.Name)),
                            "changes": {
                                name: {"from": before.get(name), "to": after.get(name)}
                                for name in normalized
                            },
                        }
                    )

                if not dry_run:
                    _safe(lambda: app.CalculateProject())

            return {"dry_run": bool(dry_run), "updated": len(results), "results": results}

        return self._run(job, timeout=600)

    def delete_tasks(
        self, task_ids: list[Any], dry_run: bool = False, project: Any = None
    ) -> dict[str, Any]:
        if not isinstance(task_ids, list) or not task_ids:
            raise InvalidInput("Provide a non-empty list of task ids.")

        def job():
            app = self._app_ref()
            proj = self._project_ref(project)
            deleted, missing = [], []
            with _quiet(app):
                for task_id in task_ids:
                    task = self._find_task(proj, task_id)
                    if task is None:
                        missing.append(task_id)
                        continue
                    record = {
                        "unique_id": _safe(lambda: int(task.UniqueID)),
                        "id": _safe(lambda: int(task.ID)),
                        "name": _safe(lambda: str(task.Name)),
                    }
                    if not dry_run:
                        task.Delete()
                    deleted.append(record)
                if not dry_run:
                    _safe(lambda: app.CalculateProject())
            return {
                "dry_run": bool(dry_run),
                "deleted": deleted,
                "missing": missing,
                "count": len(deleted),
            }

        return self._run(job, timeout=600)

    def indent_task(
        self, unique_id: Any, levels: int = 1, project: Any = None
    ) -> dict[str, Any]:
        def job():
            app = self._app_ref()
            proj = self._project_ref(project)
            task = self._find_task(proj, unique_id)
            if task is None:
                raise NotFound(f"No task with id {unique_id}.")
            with _quiet(app):
                current = int(_safe(lambda: task.OutlineLevel, 1))
                self._set_outline_level(app, task, current + int(levels))
                _safe(lambda: app.CalculateProject())
            return {
                "unique_id": _safe(lambda: int(task.UniqueID)),
                "name": _safe(lambda: str(task.Name)),
                "outline_level": _safe(lambda: int(task.OutlineLevel)),
            }

        return self._run(job)

    def move_task(
        self, unique_id: Any, after_unique_id: Any, project: Any = None
    ) -> dict[str, Any]:
        def job():
            app = self._app_ref()
            proj = self._project_ref(project)
            _safe(lambda: proj.Activate())
            task = self._find_task(proj, unique_id)
            target = self._find_task(proj, after_unique_id)
            if task is None:
                raise NotFound(f"No task with id {unique_id}.")
            if target is None:
                raise NotFound(f"No task with id {after_unique_id}.")
            with _quiet(app):
                app.SelectRow(task.ID, False)
                app.EditCut()
                target = self._find_task(proj, after_unique_id)
                if target is None:
                    raise MSProjectMCPError(
                        "The target task disappeared during the move; the schedule is "
                        "unchanged apart from the cut task now sitting on the clipboard."
                    )
                app.SelectRow(target.ID, False)
                app.EditPaste()
                _safe(lambda: app.CalculateProject())
            return {"moved": int(unique_id), "after": int(after_unique_id)}

        return self._run(job)

    # ------------------------------------------------------------------
    # Dependencies
    # ------------------------------------------------------------------

    def set_task_links(
        self,
        unique_id: Any,
        links: list[dict[str, Any]],
        replace: bool = True,
        project: Any = None,
    ) -> dict[str, Any]:
        if not isinstance(links, list):
            raise InvalidInput("links must be a list of link objects.")

        def job():
            app = self._app_ref()
            proj = self._project_ref(project)
            task = self._find_task(proj, unique_id)
            if task is None:
                raise NotFound(f"No task with id {unique_id}.")
            mpd = self._minutes_per_day(proj)

            parts: list[str] = []
            if not replace:
                existing = str(_safe(lambda: task.Predecessors, "") or "")
                parts = [part.strip() for part in existing.split(",") if part.strip()]

            resolved = []
            for link in links:
                if not isinstance(link, dict):
                    raise InvalidInput("Each link must be a JSON object.")
                predecessor_id = link.get(
                    "predecessor", link.get("unique_id", link.get("task", link.get("id")))
                )
                if predecessor_id is None:
                    raise InvalidInput("Each link needs 'predecessor' (a task id).")
                predecessor = self._find_task(proj, predecessor_id)
                if predecessor is None:
                    raise NotFound(f"Predecessor task {predecessor_id} was not found.")
                link_type = str(link.get("type", "FS")).upper()
                if link_type not in {"FS", "SS", "FF", "SF"}:
                    raise InvalidInput(
                        f"Invalid link type {link_type!r}. Use FS, SS, FF, or SF."
                    )
                lag = _format_lag(link.get("lag"), mpd)
                parts.append(f"{predecessor.ID}{link_type}{lag}")
                resolved.append(
                    {
                        "predecessor_id": _safe(lambda: int(predecessor.UniqueID)),
                        "type": link_type,
                        "lag": lag or "0d",
                    }
                )

            with _quiet(app):
                task.Predecessors = ",".join(parts)
                _safe(lambda: app.CalculateProject())

            return {
                "unique_id": _safe(lambda: int(task.UniqueID)),
                "name": _safe(lambda: str(task.Name)),
                "predecessors": str(_safe(lambda: task.Predecessors, "") or ""),
                "links": resolved,
            }

        return self._run(job)

    def get_task_links(self, unique_id: Any, project: Any = None) -> dict[str, Any]:
        def job():
            proj = self._project_ref(project)
            task = self._find_task(proj, unique_id)
            if task is None:
                raise NotFound(f"No task with id {unique_id}.")
            return {
                "unique_id": _safe(lambda: int(task.UniqueID)),
                "name": _safe(lambda: str(task.Name)),
                "predecessors": _parse_links(_safe(lambda: task.Predecessors, "")),
                "successors": _parse_links(_safe(lambda: task.Successors, "")),
            }

        return self._run(job)

    # ------------------------------------------------------------------
    # Resources and assignments
    # ------------------------------------------------------------------

    def get_resources(
        self,
        fields: Any = None,
        name_contains: str | None = None,
        limit: int | None = None,
        offset: int = 0,
        project: Any = None,
    ) -> dict[str, Any]:
        def job():
            proj = self._project_ref(project)
            mpd, mpw = self._minutes_per_day(proj), self._minutes_per_week(proj)
            names = normalize_field_list(RESOURCE_FIELDS, fields, DEFAULT_RESOURCE_FIELDS)
            cap = min(int(limit or config.MAX_ROWS), config.MAX_ROWS)
            needle = (name_contains or "").strip().lower()

            items, matched = [], 0
            for resource in proj.Resources:
                if resource is None:
                    continue
                if needle and needle not in str(_safe(lambda: resource.Name, "")).lower():
                    continue
                matched += 1
                if matched <= int(offset):
                    continue
                if len(items) >= cap:
                    continue
                items.append(self._read_fields(resource, RESOURCE_FIELDS, names, mpd, mpw))

            return {
                "project": _safe(lambda: str(proj.Name)),
                "matched": matched,
                "returned": len(items),
                "truncated": matched - int(offset) > len(items),
                "items": items,
            }

        return self._run(job, timeout=300)

    def create_resources(
        self, resources: list[dict[str, Any]], project: Any = None
    ) -> dict[str, Any]:
        if not isinstance(resources, list) or not resources:
            raise InvalidInput("Provide a non-empty list of resource objects.")

        def job():
            app = self._app_ref()
            proj = self._project_ref(project)
            mpd, mpw = self._minutes_per_day(proj), self._minutes_per_week(proj)
            created = []
            with _quiet(app):
                for item in resources:
                    if not isinstance(item, dict) or not str(item.get("name") or "").strip():
                        raise InvalidInput("Each resource needs a non-empty 'name'.")
                    resource = proj.Resources.Add(str(item["name"]).strip())
                    values = dict(item)
                    values.pop("name", None)
                    self._write_fields(
                        resource, RESOURCE_FIELDS, values, mpd, mpw, proj
                    )
                    created.append(resource)
            names = normalize_field_list(RESOURCE_FIELDS, None, DEFAULT_RESOURCE_FIELDS)
            return {
                "created": len(created),
                "items": [
                    self._read_fields(resource, RESOURCE_FIELDS, names, mpd, mpw)
                    for resource in created
                ],
            }

        return self._run(job, timeout=300)

    def update_resources(
        self,
        updates: list[dict[str, Any]],
        dry_run: bool = False,
        project: Any = None,
    ) -> dict[str, Any]:
        if not isinstance(updates, list) or not updates:
            raise InvalidInput("Provide a non-empty list of updates.")

        def job():
            app = self._app_ref()
            proj = self._project_ref(project)
            mpd, mpw = self._minutes_per_day(proj), self._minutes_per_week(proj)
            results = []
            with _quiet(app):
                for update in updates:
                    if not isinstance(update, dict):
                        raise InvalidInput("Each update must be a JSON object.")
                    selector = update.get(
                        "resource", update.get("unique_id", update.get("name"))
                    )
                    if selector is None:
                        raise InvalidInput("Each resource update needs 'resource' (name or id).")
                    resource = self._find_resource(proj, selector)
                    if resource is None:
                        raise NotFound(f"Resource {selector!r} was not found.")
                    values = update.get("set") or {
                        key: value
                        for key, value in update.items()
                        if key not in {"resource", "unique_id", "name", "id"}
                    }
                    if not values:
                        raise InvalidInput(f"No fields to change for resource {selector!r}.")

                    normalized = {}
                    for key, value in values.items():
                        name = normalize_name(key)
                        if name.endswith("_minutes"):
                            base = name[: -len("_minutes")]
                            if base in RESOURCE_FIELDS and RESOURCE_FIELDS[base].kind == "duration":
                                name = base
                        normalized[name] = value
                    names = [name for name in normalized if name in RESOURCE_FIELDS]
                    before = self._read_fields(resource, RESOURCE_FIELDS, names, mpd, mpw)
                    if dry_run:
                        results.append(
                            {
                                "resource": _safe(lambda: str(resource.Name)),
                                "changes": {
                                    name: {"from": before.get(name), "to": value}
                                    for name, value in normalized.items()
                                },
                            }
                        )
                        continue
                    self._write_fields(resource, RESOURCE_FIELDS, normalized, mpd, mpw, proj)
                    after = self._read_fields(resource, RESOURCE_FIELDS, names, mpd, mpw)
                    results.append(
                        {
                            "resource": _safe(lambda: str(resource.Name)),
                            "changes": {name: {"from": before.get(name), "to": after.get(name)}
                                        for name in names},
                        }
                    )
            return {"dry_run": bool(dry_run), "updated": len(results), "results": results}

        return self._run(job, timeout=300)

    def delete_resources(
        self, selectors: list[Any], dry_run: bool = False, project: Any = None
    ) -> dict[str, Any]:
        if not isinstance(selectors, list) or not selectors:
            raise InvalidInput("Provide a non-empty list of resource names or ids.")

        def job():
            app = self._app_ref()
            proj = self._project_ref(project)
            removed, missing = [], []
            with _quiet(app):
                for selector in selectors:
                    resource = self._find_resource(proj, selector)
                    if resource is None:
                        missing.append(selector)
                        continue
                    record = {
                        "unique_id": _safe(lambda: int(resource.UniqueID)),
                        "name": _safe(lambda: str(resource.Name)),
                    }
                    if not dry_run:
                        resource.Delete()
                    removed.append(record)
                if not dry_run:
                    _safe(lambda: app.CalculateProject())
            return {
                "dry_run": bool(dry_run),
                "deleted": removed,
                "missing": missing,
                "count": len(removed),
            }

        return self._run(job, timeout=600)

    def assign_resources(
        self,
        assignments: list[dict[str, Any]],
        create_missing: bool = True,
        project: Any = None,
    ) -> dict[str, Any]:
        if not isinstance(assignments, list) or not assignments:
            raise InvalidInput("Provide a non-empty list of assignments.")

        def job():
            app = self._app_ref()
            proj = self._project_ref(project)
            mpd, mpw = self._minutes_per_day(proj), self._minutes_per_week(proj)
            results = []
            with _quiet(app):
                for item in assignments:
                    if not isinstance(item, dict):
                        raise InvalidInput("Each assignment must be a JSON object.")
                    task_id = item.get("task", item.get("unique_id"))
                    resource_selector = item.get("resource", item.get("resource_name"))
                    if task_id is None or resource_selector is None:
                        raise InvalidInput(
                            "Each assignment needs 'task' (id) and 'resource' (name or id)."
                        )
                    task = self._find_task(proj, task_id)
                    if task is None:
                        raise NotFound(f"Task {task_id} was not found.")

                    resource = self._find_resource(proj, resource_selector)
                    if resource is None:
                        if not create_missing:
                            raise NotFound(f"Resource {resource_selector!r} was not found.")
                        resource = proj.Resources.Add(str(resource_selector).strip())

                    units = float(item.get("units", 1.0) or 1.0)
                    existing = None
                    for assignment in task.Assignments:
                        if assignment is None:
                            continue
                        if str(_safe(lambda: assignment.ResourceName, "")).lower() == str(
                            _safe(lambda: resource.Name, "")
                        ).lower():
                            existing = assignment
                            break

                    if existing is not None:
                        existing.Units = units
                        created = False
                    else:
                        # Assignments.Add takes (TaskID, ResourceID, Units) positionally.
                        # Named arguments fail, and Project.Assignments is not exposed
                        # by every build.
                        try:
                            task.Assignments.Add(task.ID, resource.ID, units)
                        except Exception as exc:
                            raise MSProjectMCPError(
                                f"Could not assign {resource.Name!r} to "
                                f"{task.Name!r}: {exc}"
                            ) from exc
                        created = True

                    results.append(
                        {
                            "task": _safe(lambda: int(task.UniqueID)),
                            "task_name": _safe(lambda: str(task.Name)),
                            "resource": _safe(lambda: str(resource.Name)),
                            "units": round(units, 4),
                            "created": created,
                        }
                    )

                _safe(lambda: app.CalculateProject())

            return {"assigned": len(results), "results": results}

        return self._run(job, timeout=600)

    def unassign_resources(
        self, items: list[dict[str, Any]], project: Any = None
    ) -> dict[str, Any]:
        if not isinstance(items, list) or not items:
            raise InvalidInput("Provide a non-empty list of assignments to remove.")

        def job():
            app = self._app_ref()
            proj = self._project_ref(project)
            removed, missing = [], []
            with _quiet(app):
                for item in items:
                    if not isinstance(item, dict):
                        raise InvalidInput("Each item must be a JSON object.")
                    task_id = item.get("task", item.get("unique_id"))
                    resource_selector = item.get("resource", item.get("resource_name"))
                    if task_id is None or resource_selector is None:
                        raise InvalidInput("Each item needs 'task' and 'resource'.")
                    task = self._find_task(proj, task_id)
                    if task is None:
                        missing.append(item)
                        continue
                    needle = str(resource_selector).strip().lower()
                    found = False
                    for assignment in task.Assignments:
                        if assignment is None:
                            continue
                        if str(_safe(lambda: assignment.ResourceName, "")).strip().lower() == needle:
                            assignment.Delete()
                            found = True
                            break
                    if found:
                        removed.append(
                            {"task": _safe(lambda: int(task.UniqueID)), "resource": needle}
                        )
                    else:
                        missing.append(item)
                _safe(lambda: app.CalculateProject())
            return {"removed": removed, "missing": missing, "count": len(removed)}

        return self._run(job, timeout=300)

    def get_assignments(
        self,
        task: Any = None,
        resource: Any = None,
        limit: int | None = None,
        project: Any = None,
    ) -> dict[str, Any]:
        def job():
            proj = self._project_ref(project)
            mpd, mpw = self._minutes_per_day(proj), self._minutes_per_week(proj)
            cap = min(int(limit or config.MAX_ROWS), config.MAX_ROWS)

            if task is not None:
                owner = self._find_task(proj, task)
                if owner is None:
                    raise NotFound(f"Task {task} was not found.")
                owners = [owner]
            else:
                # Project.Assignments is not exposed by every build; walk the tasks.
                owners = [t for t in proj.Tasks if t is not None]

            needle = ""
            if resource is not None:
                found = self._find_resource(proj, resource)
                if found is None:
                    raise NotFound(f"Resource {resource!r} was not found.")
                needle = str(_safe(lambda: found.Name, "")).strip().lower()

            items: list[dict[str, Any]] = []
            for owner in owners:
                for assignment in owner.Assignments:
                    if assignment is None:
                        continue
                    name = str(_safe(lambda: assignment.ResourceName, ""))
                    if needle and name.strip().lower() != needle:
                        continue
                    if len(items) >= cap:
                        break
                    items.append(
                        {
                            "task_id": _safe(lambda: int(assignment.TaskID)),
                            "task_unique_id": _safe(lambda: int(owner.UniqueID)),
                            "resource_id": _safe(lambda: int(assignment.ResourceID)),
                            "resource": name,
                            "units": _safe(lambda: round(float(assignment.Units), 4)),
                            "work": format_duration(
                                _safe(lambda: _as_float(assignment.Work), 0.0), mpd, mpw
                            ),
                            "actual_work": format_duration(
                                _safe(lambda: _as_float(assignment.ActualWork), 0.0),
                                mpd,
                                mpw,
                            ),
                            "remaining_work": format_duration(
                                _safe(lambda: _as_float(assignment.RemainingWork), 0.0),
                                mpd,
                                mpw,
                            ),
                            "cost": _safe(lambda: round(_as_float(assignment.Cost) or 0.0, 2)),
                            "start": format_datetime(_safe(lambda: assignment.Start)),
                            "finish": format_datetime(_safe(lambda: assignment.Finish)),
                            "percent_work_complete": _safe(
                                lambda: int(assignment.PercentWorkComplete)
                            ),
                        }
                    )
                if len(items) >= cap:
                    break
            return {"returned": len(items), "truncated": len(items) >= cap, "items": items}

        return self._run(job, timeout=300)

    # ------------------------------------------------------------------
    # Calendars
    # ------------------------------------------------------------------

    def get_calendars(self, project: Any = None) -> dict[str, Any]:
        def job():
            proj = self._project_ref(project)
            calendars = []
            collection = _safe(lambda: proj.BaseCalendars)
            count = _safe(lambda: int(collection.Count), 0) if collection is not None else 0
            for index in range(1, count + 1):
                calendar = collection(index)
                calendars.append(
                    {
                        "index": index,
                        "name": _safe(lambda: str(calendar.Name)),
                        "guid": _safe(lambda: str(calendar.Guid)),
                    }
                )
            return {
                "project": _safe(lambda: str(proj.Name)),
                "project_calendar": _safe(lambda: str(proj.Calendar.Name)),
                "count": count,
                "calendars": calendars,
            }

        return self._run(job)

    def create_calendar(
        self, name: str, copy_from: str = "Standard", project: Any = None
    ) -> dict[str, Any]:
        if not str(name or "").strip():
            raise InvalidInput("A calendar name is required.")

        def job():
            app = self._app_ref()
            proj = self._project_ref(project)
            source = self._calendar(proj, copy_from)
            if _safe(lambda: str(source.Name)) == str(name):
                raise InvalidInput(f"Calendar {name!r} already exists.")
            with _quiet(app):
                try:
                    app.BaseCalendarCreate(Name=str(name), FromName=str(source.Name))
                except Exception:
                    app.BaseCalendarCreate(str(name), str(source.Name))
            return {"created": str(name), "copied_from": str(source.Name)}

        return self._run(job)

    def add_calendar_exception(
        self,
        calendar: str,
        name: str,
        start: str,
        finish: str | None = None,
        working: bool = False,
        project: Any = None,
    ) -> dict[str, Any]:
        def job():
            app = self._app_ref()
            proj = self._project_ref(project)
            target = self._calendar(proj, calendar)
            start_dt = parse_datetime(start)
            finish_dt = parse_datetime(finish) if finish else start_dt
            if finish_dt < start_dt:
                raise InvalidInput("The exception finish date is before its start date.")
            with _quiet(app):
                # pjCalendarExceptionDaily = 1
                exception = target.Exceptions.Add(1, start_dt, finish_dt, str(name))
                if working:
                    _safe(lambda: setattr(exception, "Working", True))
            return {
                "calendar": str(_safe(lambda: target.Name)),
                "name": str(name),
                "start": format_datetime(start_dt),
                "finish": format_datetime(finish_dt),
                "working": bool(working),
            }

        return self._run(job)

    def list_calendar_exceptions(
        self, calendar: str, project: Any = None
    ) -> dict[str, Any]:
        def job():
            proj = self._project_ref(project)
            target = self._calendar(proj, calendar)
            exceptions = []
            for exception in target.Exceptions:
                if exception is None:
                    continue
                exceptions.append(
                    {
                        "index": _safe(lambda: int(exception.Index)),
                        "name": _safe(lambda: str(exception.Name)),
                        "start": format_datetime(_safe(lambda: exception.Start)),
                        "finish": format_datetime(_safe(lambda: exception.Finish)),
                    }
                )
            return {
                "calendar": str(_safe(lambda: target.Name)),
                "count": len(exceptions),
                "exceptions": exceptions,
            }

        return self._run(job)

    def delete_calendar_exception(
        self, calendar: str, index: int, project: Any = None
    ) -> dict[str, Any]:
        def job():
            app = self._app_ref()
            proj = self._project_ref(project)
            target = self._calendar(proj, calendar)
            with _quiet(app):
                target.Exceptions(int(index)).Delete()
            return {"calendar": str(_safe(lambda: target.Name)), "deleted_index": int(index)}

        return self._run(job)

    def set_working_hours(
        self,
        calendar: str,
        day: str,
        shifts: list[list[str]],
        project: Any = None,
    ) -> dict[str, Any]:
        key = str(day).strip().lower()
        day_index = _WEEKDAYS.get(key)
        if day_index is None:
            if key.isdigit() and 1 <= int(key) <= 7:
                day_index = int(key)
            else:
                raise InvalidInput(
                    f"Unknown day {day!r}. Use sunday..saturday or 1..7 (1=Sunday)."
                )
        if not isinstance(shifts, list):
            raise InvalidInput("shifts must be a list of [start, finish] pairs.")

        def job():
            app = self._app_ref()
            proj = self._project_ref(project)
            target = self._calendar(proj, calendar)
            weekday = target.WeekDays(day_index)
            with _quiet(app):
                if not shifts:
                    weekday.Working = False
                else:
                    weekday.Working = True
                    for position in range(1, 6):
                        shift = _safe(lambda: getattr(weekday, f"Shift{position}"))
                        if shift is None:
                            continue
                        if position <= len(shifts):
                            pair = shifts[position - 1]
                            if not isinstance(pair, (list, tuple)) or len(pair) != 2:
                                raise InvalidInput(
                                    "Each shift must be a [start, finish] pair, e.g. [\"08:00\", \"12:00\"]."
                                )
                            shift.Start = str(pair[0])
                            shift.Finish = str(pair[1])
                        else:
                            _safe(lambda: setattr(shift, "Start", ""))
                            _safe(lambda: setattr(shift, "Finish", ""))
            return {
                "calendar": str(_safe(lambda: target.Name)),
                "day": day_index,
                "working": bool(shifts),
                "shifts": shifts,
            }

        return self._run(job)

    # ------------------------------------------------------------------
    # Scheduling and analysis
    # ------------------------------------------------------------------

    def calculate(self, project: Any = None) -> dict[str, Any]:
        def job():
            app = self._app_ref()
            proj = self._project_ref(project)
            with _quiet(app):
                self._activate(app, proj)
                app.CalculateProject()
            return {
                "project": _safe(lambda: str(proj.Name)),
                "finish": format_datetime(_safe(lambda: proj.ProjectFinish)),
            }

        return self._run(job, timeout=600)

    def level_resources(self, project: Any = None) -> dict[str, Any]:
        def job():
            app = self._app_ref()
            proj = self._project_ref(project)
            with _quiet(app):
                self._activate(app, proj)
                app.LevelNow()
            return {
                "project": _safe(lambda: str(proj.Name)),
                "finish": format_datetime(_safe(lambda: proj.ProjectFinish)),
            }

        return self._run(job, timeout=900)

    def save_baseline(self, number: int = 0, project: Any = None) -> dict[str, Any]:
        if not 0 <= int(number) <= 10:
            raise InvalidInput("Baseline number must be 0-10.")

        def job():
            app = self._app_ref()
            proj = self._project_ref(project)
            tasks = _safe(lambda: int(proj.Tasks.Count), 0)
            with _quiet(app):
                self._activate(app, proj)
                app.BaselineSave(All=True, Copy=int(number), Into=int(number))
            return {"saved": int(number), "tasks": tasks}

        return self._run(job, timeout=600)

    def clear_baseline(self, number: int = 0, project: Any = None) -> dict[str, Any]:
        if not 0 <= int(number) <= 10:
            raise InvalidInput("Baseline number must be 0-10.")

        def job():
            app = self._app_ref()
            proj = self._project_ref(project)
            with _quiet(app):
                self._activate(app, proj)
                app.BaselineClear(All=True, From=int(number))
            return {"cleared": int(number)}

        return self._run(job, timeout=600)

    @staticmethod
    def _baseline_attr(number: int, suffix: str) -> str:
        prefix = "Baseline" if int(number) == 0 else f"Baseline{int(number)}"
        return f"{prefix}{suffix}"

    def baseline_variance(self, number: int = 0, project: Any = None) -> dict[str, Any]:
        if not 0 <= int(number) <= 10:
            raise InvalidInput("Baseline number must be 0-10.")

        def job():
            proj = self._project_ref(project)
            mpd, mpw = self._minutes_per_day(proj), self._minutes_per_week(proj)
            start_attr = self._baseline_attr(number, "Start")
            finish_attr = self._baseline_attr(number, "Finish")
            duration_attr = self._baseline_attr(number, "Duration")
            work_attr = self._baseline_attr(number, "Work")
            cost_attr = self._baseline_attr(number, "Cost")

            rows = []
            for task in proj.Tasks:
                if task is None or _safe(lambda: bool(task.Summary), False):
                    continue
                baseline_finish = _safe(lambda: getattr(task, finish_attr))
                if baseline_finish is None:
                    continue
                finish = _safe(lambda: task.Finish)
                finish_variance = 0.0
                if finish is not None and baseline_finish is not None:
                    finish_variance = (
                        parse_datetime(finish) - parse_datetime(baseline_finish)
                    ).total_seconds() / 86400.0
                rows.append(
                    {
                        "unique_id": _safe(lambda: int(task.UniqueID)),
                        "id": _safe(lambda: int(task.ID)),
                        "name": _safe(lambda: str(task.Name)),
                        "baseline_start": format_datetime(
                            _safe(lambda: getattr(task, start_attr))
                        ),
                        "start": format_datetime(_safe(lambda: task.Start)),
                        "baseline_finish": format_datetime(baseline_finish),
                        "finish": format_datetime(finish),
                        "finish_variance_days": round(finish_variance, 2),
                        "baseline_duration": format_duration(
                            _safe(lambda: float(getattr(task, duration_attr)), 0.0), mpd, mpw
                        ),
                        "duration": format_duration(
                            _safe(lambda: float(task.Duration), 0.0), mpd, mpw
                        ),
                        "baseline_work": format_duration(
                            _safe(lambda: float(getattr(task, work_attr)), 0.0), mpd, mpw
                        ),
                        "work": format_duration(
                            _safe(lambda: float(task.Work), 0.0), mpd, mpw
                        ),
                        "baseline_cost": _safe(lambda: round(float(getattr(task, cost_attr)), 2)),
                        "cost": _safe(lambda: round(float(task.Cost), 2)),
                    }
                )

            cost_variance = sum(
                (row["cost"] or 0.0) - (row["baseline_cost"] or 0.0) for row in rows
            )
            late = sum(1 for row in rows if row["finish_variance_days"] > 0.01)
            return {
                "baseline": int(number),
                "tasks_compared": len(rows),
                "tasks_late": late,
                "cost_variance": round(cost_variance, 2),
                "items": rows[: config.MAX_ROWS],
                "truncated": len(rows) > config.MAX_ROWS,
            }

        return self._run(job, timeout=600)

    def earned_value(self, project: Any = None) -> dict[str, Any]:
        def job():
            proj = self._project_ref(project)
            rows, totals = [], {"bcws": 0.0, "bcwp": 0.0, "acwp": 0.0}
            for task in proj.Tasks:
                if task is None or _safe(lambda: bool(task.Summary), False):
                    continue
                row = {
                    "unique_id": _safe(lambda: int(task.UniqueID)),
                    "name": _safe(lambda: str(task.Name)),
                    "bcws": _safe(lambda: round(float(task.BCWS), 2), 0.0),
                    "bcwp": _safe(lambda: round(float(task.BCWP), 2), 0.0),
                    "acwp": _safe(lambda: round(float(task.ACWP), 2), 0.0),
                    "sv": _safe(lambda: round(float(task.SV), 2), 0.0),
                    "cv": _safe(lambda: round(float(task.CV), 2), 0.0),
                    "spi": _safe(lambda: round(float(task.SPI), 3)),
                    "cpi": _safe(lambda: round(float(task.CPI), 3)),
                }
                totals["bcws"] += row["bcws"]
                totals["bcwp"] += row["bcwp"]
                totals["acwp"] += row["acwp"]
                rows.append(row)
            spi = (totals["bcwp"] / totals["bcws"]) if totals["bcws"] else None
            cpi = (totals["bcwp"] / totals["acwp"]) if totals["acwp"] else None
            return {
                "totals": {key: round(value, 2) for key, value in totals.items()},
                "spi": round(spi, 3) if spi is not None else None,
                "cpi": round(cpi, 3) if cpi is not None else None,
                "items": rows[: config.MAX_ROWS],
                "truncated": len(rows) > config.MAX_ROWS,
            }

        return self._run(job, timeout=600)

    def critical_path(self, project: Any = None) -> dict[str, Any]:
        def job():
            proj = self._project_ref(project)
            mpd, mpw = self._minutes_per_day(proj), self._minutes_per_week(proj)
            tasks = []
            for task in proj.Tasks:
                if task is None or _safe(lambda: bool(task.Summary), False):
                    continue
                if not _safe(lambda: bool(task.Critical), False):
                    continue
                tasks.append(
                    {
                        "unique_id": _safe(lambda: int(task.UniqueID)),
                        "id": _safe(lambda: int(task.ID)),
                        "name": _safe(lambda: str(task.Name)),
                        "start": format_datetime(_safe(lambda: task.Start)),
                        "finish": format_datetime(_safe(lambda: task.Finish)),
                        "duration": format_duration(
                            _safe(lambda: float(task.Duration), 0.0), mpd, mpw
                        ),
                        "total_slack": format_duration(
                            _safe(lambda: float(task.TotalSlack), 0.0), mpd, mpw
                        ),
                        "predecessors": str(_safe(lambda: task.Predecessors, "") or ""),
                    }
                )
            tasks.sort(key=lambda row: (row["start"] or "", row["id"]))
            return {
                "project": _safe(lambda: str(proj.Name)),
                "critical_tasks": len(tasks),
                "project_finish": format_datetime(_safe(lambda: proj.ProjectFinish)),
                "items": tasks[: config.MAX_ROWS],
                "truncated": len(tasks) > config.MAX_ROWS,
            }

        return self._run(job, timeout=600)

    def check_schedule(self, project: Any = None) -> dict[str, Any]:
        def job():
            proj = self._project_ref(project)
            mpd, mpw = self._minutes_per_day(proj), self._minutes_per_week(proj)
            today = dt.date.today()
            overdue, negative_slack, hard_constraints, unassigned, no_links = [], [], [], [], []
            task_count = 0

            for task in proj.Tasks:
                if task is None:
                    continue
                task_count += 1
                summary = _safe(lambda: bool(task.Summary), False)
                if summary:
                    continue
                record = {
                    "unique_id": _safe(lambda: int(task.UniqueID)),
                    "id": _safe(lambda: int(task.ID)),
                    "name": _safe(lambda: str(task.Name)),
                }
                percent = _safe(lambda: int(task.PercentComplete), 0)
                finish = _safe(lambda: task.Finish)
                if finish is not None and percent < 100:
                    if parse_datetime(finish).date() < today:
                        overdue.append({**record, "finish": format_datetime(finish), "percent_complete": percent})

                slack = _safe(lambda: float(task.TotalSlack), 0.0)
                if slack < 0:
                    negative_slack.append(
                        {**record, "total_slack": format_duration(slack, mpd, mpw)}
                    )

                constraint = _safe(lambda: int(task.ConstraintType), 0)
                if constraint not in (0, 1):
                    hard_constraints.append(
                        {
                            **record,
                            "constraint_type": enum_name(CONSTRAINT_TYPES, constraint),
                            "constraint_date": format_datetime(
                                _safe(lambda: task.ConstraintDate)
                            ),
                        }
                    )

                if not str(_safe(lambda: task.ResourceNames, "") or "").strip():
                    unassigned.append(record)
                if not str(_safe(lambda: task.Predecessors, "") or "").strip():
                    no_links.append(record)

            issues = (
                len(overdue)
                + len(negative_slack)
                + len(hard_constraints)
                + len(overdue)
            )
            return {
                "project": _safe(lambda: str(proj.Name)),
                "task_count": task_count,
                "issue_count": issues,
                "overdue": overdue[:200],
                "negative_slack": negative_slack[:200],
                "hard_constraints": hard_constraints[:200],
                "tasks_without_resources": unassigned[:200],
                "tasks_without_predecessors": no_links[:200],
            }

        return self._run(job, timeout=600)

    def simulate_delay(
        self, unique_id: Any, days: float, project: Any = None
    ) -> dict[str, Any]:
        """Apply a Start No Earlier Than constraint, measure the impact, then revert."""

        def job():
            app = self._app_ref()
            proj = self._project_ref(project)
            self._activate(app, proj)
            task = self._find_task(proj, unique_id)
            if task is None:
                raise NotFound(f"No task with id {unique_id}.")

            old_type = _safe(lambda: int(task.ConstraintType), 0)
            old_date = _safe(lambda: task.ConstraintDate)
            old_finish = _safe(lambda: proj.ProjectFinish)
            old_start = _safe(lambda: task.Start)
            if old_start is None:
                raise MSProjectMCPError("This task has no start date to delay.")

            new_start = parse_datetime(old_start) + dt.timedelta(days=float(days))
            try:
                with _quiet(app):
                    # pjSNET = 4
                    task.ConstraintType = 4
                    task.ConstraintDate = new_start
                    app.CalculateProject()

                new_finish = _safe(lambda: proj.ProjectFinish)
                critical = [
                    str(_safe(lambda: t.Name, ""))
                    for t in proj.Tasks
                    if t is not None
                    and not _safe(lambda: bool(t.Summary), False)
                    and _safe(lambda: bool(t.Critical), False)
                ]
                shift = 0.0
                if old_finish is not None and new_finish is not None:
                    shift = (
                        parse_datetime(new_finish) - parse_datetime(old_finish)
                    ).total_seconds() / 86400.0
                result = {
                    "task": _safe(lambda: str(task.Name)),
                    "delay_days": float(days),
                    "old_start": format_datetime(old_start),
                    "new_start": format_datetime(new_start),
                    "old_project_finish": format_datetime(old_finish),
                    "new_project_finish": format_datetime(new_finish),
                    "project_finish_shift_days": round(shift, 2),
                    "critical_task_count": len(critical),
                    "critical_tasks": critical[:200],
                }
            finally:
                with _quiet(app):
                    _safe(lambda: setattr(task, "ConstraintType", old_type))
                    _safe(lambda: setattr(task, "ConstraintDate", old_date))
                    _safe(lambda: app.CalculateProject())
            result["reverted"] = True
            return result

        return self._run(job, timeout=900)

    def timephased(
        self,
        unique_id: Any,
        start: str,
        finish: str,
        data_type: str = "work",
        timescale: str = "weekly",
        project: Any = None,
    ) -> dict[str, Any]:
        scales = {"daily": 3, "weekly": 4, "monthly": 5}
        types = {
            "work": 1,
            "cost": 2,
            "actual_work": 3,
            "actual_cost": 4,
            "remaining_work": 9,
            "baseline_work": 22,
            "baseline_cost": 23,
        }
        scale_key = str(timescale).lower()
        type_key = str(data_type).lower()
        if scale_key not in scales:
            raise InvalidInput(f"timescale must be one of {sorted(scales)}.")
        if type_key not in types:
            raise InvalidInput(f"data_type must be one of {sorted(types)}.")

        def job():
            proj = self._project_ref(project)
            task = self._find_task(proj, unique_id)
            if task is None:
                raise NotFound(f"No task with id {unique_id}.")
            start_dt = parse_datetime(start)
            finish_dt = parse_datetime(finish)
            if finish_dt < start_dt:
                raise InvalidInput("finish must be on or after start.")
            series = task.TimeScaleData(start_dt, finish_dt, types[type_key], scales[scale_key])
            points = []
            for point in series:
                value = _safe(lambda: float(point.Value), 0.0)
                points.append(
                    {
                        "start": format_datetime(_safe(lambda: point.StartDate)),
                        "finish": format_datetime(_safe(lambda: point.EndDate)),
                        "value": round(value, 4),
                    }
                )
            return {
                "task": _safe(lambda: str(task.Name)),
                "data_type": type_key,
                "timescale": scale_key,
                "points": points,
            }

        return self._run(job, timeout=600)

    # ------------------------------------------------------------------
    # Import / export
    # ------------------------------------------------------------------

    def export_csv(
        self,
        path: str,
        fields: Any = None,
        project: Any = None,
    ) -> dict[str, Any]:
        target = os.path.abspath(os.path.expanduser(str(path)))

        def job():
            proj = self._project_ref(project)
            mpd, mpw = self._minutes_per_day(proj), self._minutes_per_week(proj)
            names = normalize_field_list(TASK_FIELDS, fields, DEFAULT_TASK_FIELDS)
            rows = [
                self._read_fields(task, TASK_FIELDS, names, mpd, mpw)
                for task in proj.Tasks
                if task is not None
            ]
            columns: list[str] = []
            for row in rows:
                for key in row:
                    if key not in columns:
                        columns.append(key)
            directory = os.path.dirname(target)
            if directory:
                os.makedirs(directory, exist_ok=True)
            with open(target, "w", newline="", encoding="utf-8-sig") as handle:
                writer = csv.DictWriter(handle, fieldnames=columns)
                writer.writeheader()
                writer.writerows(rows)
            return {"exported": target, "rows": len(rows), "columns": columns}

        return self._run(job, timeout=300)

    def snapshot(
        self,
        path: str | None = None,
        limit: int | None = None,
        project: Any = None,
    ) -> dict[str, Any]:
        def job():
            proj = self._project_ref(project)
            mpd, mpw = self._minutes_per_day(proj), self._minutes_per_week(proj)
            cap = int(limit or config.MAX_ROWS)
            task_fields = normalize_field_list(
                TASK_FIELDS,
                [
                    "name", "outline_level", "summary", "milestone", "duration",
                    "start", "finish", "percent_complete", "critical", "predecessors",
                    "resource_names", "constraint_type", "deadline", "notes",
                    "work", "cost",
                ],
                DEFAULT_TASK_FIELDS,
            )
            tasks = [
                self._read_fields(task, TASK_FIELDS, task_fields, mpd, mpw)
                for task in proj.Tasks
                if task is not None
            ][:cap]
            resources = [
                self._read_fields(
                    resource,
                    RESOURCE_FIELDS,
                    normalize_field_list(RESOURCE_FIELDS, None, DEFAULT_RESOURCE_FIELDS),
                    mpd,
                    mpw,
                )
                for resource in proj.Resources
                if resource is not None
            ][:cap]
            data = {
                "project": self._project_summary(proj),
                "tasks": tasks,
                "resources": resources,
                "task_count": _safe(lambda: int(proj.Tasks.Count), 0),
                "truncated": _safe(lambda: int(proj.Tasks.Count), 0) > len(tasks),
            }
            if path:
                target = os.path.abspath(os.path.expanduser(str(path)))
                directory = os.path.dirname(target)
                if directory:
                    os.makedirs(directory, exist_ok=True)
                with open(target, "w", encoding="utf-8") as handle:
                    json.dump(data, handle, indent=2, default=str)
                return {"written_to": target, "tasks": len(tasks), "resources": len(resources)}
            return data

        return self._run(job, timeout=300)

    def export_pdf(self, path: str, project: Any = None) -> dict[str, Any]:
        """Export the active view to PDF.

        Project's PDF exporter is Application.DocumentExport(Filename, FileType);
        FileType 0 is pjPDF (1 is pjXPS). Project.ExportAsFixedFormat exists but
        rejects every argument combination tried on Project 16.
        """
        target = os.path.abspath(os.path.expanduser(str(path)))

        def job():
            app = self._app_ref()
            proj = self._project_ref(project)
            directory = os.path.dirname(target)
            if directory:
                os.makedirs(directory, exist_ok=True)
            with _quiet(app):
                self._activate(app, proj)
                app.DocumentExport(target, 0)
            if not os.path.exists(target):
                raise MSProjectMCPError(
                    "Project reported success but no PDF was written to "
                    f"{target}. Save the project first, then retry."
                )
            return {
                "exported": target,
                "format": "pdf",
                "bytes": os.path.getsize(target),
            }

        return self._run(job, timeout=900)

    def import_xml(self, path: str) -> dict[str, Any]:
        target = os.path.abspath(os.path.expanduser(str(path)))
        if not os.path.exists(target):
            raise NotFound(f"File not found: {target}")
        return self.open_project(target)

    # ------------------------------------------------------------------
    # Escape hatches
    # ------------------------------------------------------------------

    def run_macro(self, name: str, args: list[Any] | None = None) -> dict[str, Any]:
        macro = str(name or "").strip()
        if not macro:
            raise InvalidInput("A macro name is required, e.g. 'Module1.MyMacro'.")

        def job():
            app = self._app_ref()
            self._project_ref()
            call_args = list(args or [])
            with _quiet(app):
                try:
                    result = app.Run(macro, *call_args) if call_args else app.Run(macro)
                except Exception as exc:
                    # Project is happiest with the bare procedure name; retry when
                    # the caller passed "Module1.Macro" or "Project.Module.Macro".
                    if "." not in macro:
                        raise
                    result = (
                        app.Run(macro.rsplit(".", 1)[-1], *call_args)
                        if call_args
                        else app.Run(macro.rsplit(".", 1)[-1])
                    )
            return {"macro": macro, "result": _jsonable(result)}

        return self._run(job, timeout=900)

    def run_vba(self, code: str) -> dict[str, Any]:
        from .vba import execute_vba

        if not config.ALLOW_VBA:
            raise MSProjectMCPError(
                "VBA execution is disabled. Set MSPROJECT_MCP_ALLOW_VBA=1 in the server "
                "environment and enable 'Trust access to the VBA project object model' "
                "in Project > Options > Trust Center > Macro Settings."
            )
        return self._run(lambda: execute_vba(self._app_ref(), self._project_ref(), code), timeout=900)

    # ------------------------------------------------------------------
    # Introspection
    # ------------------------------------------------------------------

    @staticmethod
    def field_catalog(entity: str = "task") -> dict[str, Any]:
        from .fields import catalog

        key = str(entity or "task").strip().lower()
        if key == "task":
            data = catalog(TASK_FIELDS)
            data["defaults"] = DEFAULT_TASK_FIELDS
        elif key == "resource":
            data = catalog(RESOURCE_FIELDS)
            data["defaults"] = DEFAULT_RESOURCE_FIELDS
        elif key == "project":
            data = catalog(PROJECT_FIELDS)
            data["defaults"] = [field.name for field in PROJECT_FIELDS.values()]
        else:
            raise InvalidInput("entity must be 'task', 'resource', or 'project'.")
        data["entity"] = key
        return data


def _jsonable(value: Any) -> Any:
    if value is None or isinstance(value, (bool, int, float, str)):
        return value
    if isinstance(value, (dt.datetime, dt.date)):
        return format_datetime(value)
    if isinstance(value, (list, tuple)):
        return [_jsonable(item) for item in value]
    try:
        return str(value)
    except Exception:
        return None
