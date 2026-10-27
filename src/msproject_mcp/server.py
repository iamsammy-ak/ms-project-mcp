"""MCP tool surface. Every tool is a thin wrapper around MSProjectClient."""

from __future__ import annotations

import argparse
import logging
import os
import sys
from typing import Any

from mcp.server.mcpserver import MCPServer
from mcp.types import ToolAnnotations

from .core import MSProjectClient

READ = ToolAnnotations(read_only_hint=True, open_world_hint=False)
WRITE = ToolAnnotations(
    read_only_hint=False, destructive_hint=False, idempotent_hint=False, open_world_hint=False
)
DESTRUCTIVE = ToolAnnotations(
    read_only_hint=False, destructive_hint=True, idempotent_hint=False, open_world_hint=False
)

INSTRUCTIONS = """
This server drives the Microsoft Project desktop application running on this
Windows machine through COM automation.

Operating notes:
- Call msproject_status first if you are unsure whether Project is running.
  Read-only tools never start Project; create_project or open_project does.
- Ids: task and resource tools accept a UniqueID (stable across edits) or the
  visible row ID. UniqueID is matched first. Every task row returns both.
- Durations accept "5d", "12h", "480m", or a raw minute count. Read results
  include both a formatted value and a <field>_minutes number.
- Percentages are 0-100 (a 0.0-1.0 fraction is also accepted).
- Bulk and multi-field writes accept a `fields` / `set` object keyed by field
  name; call get_field_catalog to see every valid name for tasks, resources,
  and the project.
- Prefer one create_tasks call with an `indent` on each item over many single
  calls when building a WBS; it is far faster.
- Every bulk mutation supports dry_run. Use it before large edits.
- write tools return what actually changed, read back from Project.
- The project is recalculated after mutations, but call calculate_project if
  you changed links or constraints from an external tool.

If a task cannot be expressed with a typed tool, run_vba is available when the
server was started with MSPROJECT_MCP_ALLOW_VBA=1 and Project trusts VBA project
access. Assign your result to MspMcpEval, e.g.
    Dim t: Set t = ActiveProject.Tasks(3): MspMcpEval = t.Name
"""

mcp = MCPServer(
    name="microsoft-project",
    title="Microsoft Project",
    description="Read and edit the desktop Microsoft Project application via COM.",
    instructions=INSTRUCTIONS.strip(),
    version="0.1.0",
)

client = MSProjectClient()


# ---------------------------------------------------------------------------
# Session and project lifecycle
# ---------------------------------------------------------------------------


@mcp.tool(annotations=READ)
def msproject_status() -> dict[str, Any]:
    """Report whether Microsoft Project is running, which files are open, focused project, and whether the VBA escape hatch is enabled."""
    return client.status()


@mcp.tool(annotations=READ)
def list_open_projects() -> dict[str, Any]:
    """List every project file currently open in Microsoft Project, with task/resource counts and which one is active."""
    return client.list_projects()


@mcp.tool(annotations=WRITE)
def use_project(selector: str) -> dict[str, Any]:
    """Switch the active project. `selector` is a 1-based index, a name, name fragment, or file path."""
    return client.use_project(selector)


@mcp.tool(annotations=WRITE)
def create_project(title: str | None = None, start: str | None = None) -> dict[str, Any]:
    """Create a new blank project in Project and make it the working project. `start` is ISO, e.g. 2026-01-05."""
    return client.create_project(title=title, start=start)


@mcp.tool(annotations=WRITE)
def open_project(path: str, read_only: bool = False) -> dict[str, Any]:
    """Open an .mpp, .xml, or .mpt file on this PC and make it the working project."""
    return client.open_project(path, read_only=read_only)


@mcp.tool(annotations=WRITE)
def save_project() -> dict[str, Any]:
    """Save the working project in place (Project's Ctrl+S)."""
    return client.save_project()


@mcp.tool(annotations=DESTRUCTIVE)
def save_project_as(path: str, format: str = "mpp") -> dict[str, Any]:
    """Save the working project to a new path. format: mpp (default) or mpt (template).

    Like Project's Save As, this rebinds the open project to the new file.
    For spreadsheets use export_csv; for a shareable document use export_pdf.
    """
    return client.save_project_as(path, fmt=format)


@mcp.tool(annotations=DESTRUCTIVE)
def close_project(save: bool = False, project: str | None = None) -> dict[str, Any]:
    """Close the working project (or `project`). Discards unsaved changes unless save=true."""
    return client.close_project(save=save, project=project)


@mcp.tool(annotations=DESTRUCTIVE)
def quit_msproject(save: bool = False) -> dict[str, Any]:
    """Quit the Microsoft Project application entirely. Use only when asked to shut it down."""
    return client.quit_application(save=save)


@mcp.tool(annotations=WRITE)
def set_project_window_visible(visible: bool = True) -> dict[str, Any]:
    """Show or hide the Project window while working."""
    return client.set_visible(visible)


# ---------------------------------------------------------------------------
# Project properties and introspection
# ---------------------------------------------------------------------------


@mcp.tool(annotations=READ)
def get_project(fields: list[str] | None = None, project: str | None = None) -> dict[str, Any]:
    """Read project metadata (title, start/finish, status date, calendar, minutes/day) plus any requested PROJECT fields."""
    return client.get_project(fields=fields, project=project)


@mcp.tool(annotations=WRITE)
def update_project(values: dict[str, Any], project: str | None = None) -> dict[str, Any]:
    """Update project-level properties, e.g. {"title": "Launch", "status_date": "2026-06-30", "calendar": "24 Hours"}."""
    return client.update_project(values, project=project)


@mcp.tool(annotations=READ)
def get_field_catalog(entity: str = "task") -> dict[str, Any]:
    """List every valid field name for entity='task', 'resource', or 'project', split into writable and read-only, plus the default field set."""
    return MSProjectClient.field_catalog(entity)


# ---------------------------------------------------------------------------
# Tasks
# ---------------------------------------------------------------------------


@mcp.tool(annotations=READ)
def get_tasks(
    fields: list[str] | None = None,
    name_contains: str | None = None,
    outline_level: int | None = None,
    include_summary: bool = True,
    milestones_only: bool = False,
    limit: int | None = None,
    offset: int = 0,
    project: str | None = None,
) -> dict[str, Any]:
    """List tasks with optional filters. Rows are capped (default 1000); the response reports total/matched/truncated.

    Default fields: unique_id, id, name, outline_level, summary, milestone, duration,
    start, finish, percent_complete, critical, predecessors, resource_names.
    """
    return client.get_tasks(
        fields=fields,
        name_contains=name_contains,
        outline_level=outline_level,
        include_summary=include_summary,
        milestones_only=milestones_only,
        limit=limit,
        offset=offset,
        project=project,
    )


@mcp.tool(annotations=READ)
def get_task(unique_id: int, fields: list[str] | None = None, project: str | None = None) -> dict[str, Any]:
    """Read one task by UniqueID (or visible row ID), optionally a custom field list."""
    return client.get_task(unique_id, fields=fields, project=project)


@mcp.tool(annotations=WRITE)
def create_tasks(tasks: list[dict[str, Any]], project: str | None = None) -> dict[str, Any]:
    """Append tasks in order and return the created rows.

    Each object needs "name" plus any writable task field, e.g.
      {"name": "Design", "duration": "10d", "start": "2026-01-05"}
      {"name": "Wireframes", "duration": "5d", "indent": 1}
    `indent` (0-10) nests the row under the previous task. A summary row is
    created automatically when a later task is indented beneath it.
    Tasks are auto-scheduled by default so links and durations recalculate the
    plan; pass "manual": true to pin a task's dates instead.
    """
    return client.create_tasks(tasks, project=project)


@mcp.tool(annotations=WRITE)
def update_tasks(
    updates: list[dict[str, Any]],
    dry_run: bool = False,
    project: str | None = None,
) -> dict[str, Any]:
    """Change fields on many tasks at once.

    Each item: {"unique_id": 12, "set": {"duration": "8d", "percent_complete": 40, "text1": "Red"}}
    Returns before/after values. Set dry_run=true to preview without writing.
    """
    return client.update_tasks(updates, dry_run=dry_run, project=project)


@mcp.tool(annotations=DESTRUCTIVE)
def delete_tasks(
    unique_ids: list[int], dry_run: bool = False, project: str | None = None
) -> dict[str, Any]:
    """Delete tasks by id. Deleting a summary removes its subtasks. dry_run=true previews the affected rows."""
    return client.delete_tasks(unique_ids, dry_run=dry_run, project=project)


@mcp.tool(annotations=WRITE)
def indent_task(unique_id: int, levels: int = 1, project: str | None = None) -> dict[str, Any]:
    """Indent (positive levels) or outdent (negative) a task, changing its WBS outline level."""
    return client.indent_task(unique_id, levels=levels, project=project)


@mcp.tool(annotations=DESTRUCTIVE)
def move_task(unique_id: int, after_unique_id: int, project: str | None = None) -> dict[str, Any]:
    """Move a task row so it sits immediately after another task. Uses Project's cut/paste, which touches the clipboard."""
    return client.move_task(unique_id, after_unique_id, project=project)


@mcp.tool(annotations=WRITE)
def set_task_links(
    unique_id: int,
    links: list[dict[str, Any]],
    replace: bool = True,
    project: str | None = None,
) -> dict[str, Any]:
    """Set predecessor links on a task.

    Each link: {"predecessor": 4, "type": "FS", "lag": "2d"}
    type is FS, SS, FF, or SF; lag accepts "2d", "-1d", or minutes. replace=false
    appends to the existing predecessor list.
    """
    return client.set_task_links(unique_id, links, replace=replace, project=project)


@mcp.tool(annotations=READ)
def get_task_links(unique_id: int, project: str | None = None) -> dict[str, Any]:
    """Return a task's parsed predecessors and successors (id, link type, lag)."""
    return client.get_task_links(unique_id, project=project)


# ---------------------------------------------------------------------------
# Resources and assignments
# ---------------------------------------------------------------------------


@mcp.tool(annotations=READ)
def get_resources(
    fields: list[str] | None = None,
    name_contains: str | None = None,
    limit: int | None = None,
    offset: int = 0,
    project: str | None = None,
) -> dict[str, Any]:
    """List the resource pool with optional name filter and paging."""
    return client.get_resources(
        fields=fields,
        name_contains=name_contains,
        limit=limit,
        offset=offset,
        project=project,
    )


@mcp.tool(annotations=WRITE)
def create_resources(resources: list[dict[str, Any]], project: str | None = None) -> dict[str, Any]:
    """Add resources to the pool.

    Each object needs "name"; type is work (default), material, or cost.
    Example: {"name": "Alice", "max_units": 1.0, "standard_rate": 85, "group": "Eng"}
    """
    return client.create_resources(resources, project=project)


@mcp.tool(annotations=WRITE)
def update_resources(
    updates: list[dict[str, Any]],
    dry_run: bool = False,
    project: str | None = None,
) -> dict[str, Any]:
    """Change resource properties.

    Each item: {"resource": "Alice", "set": {"max_units": 0.5, "standard_rate": 95}}
    """
    return client.update_resources(updates, dry_run=dry_run, project=project)


@mcp.tool(annotations=DESTRUCTIVE)
def delete_resources(
    resources: list[str], dry_run: bool = False, project: str | None = None
) -> dict[str, Any]:
    """Delete resources by name or id; their assignments are removed too."""
    return client.delete_resources(resources, dry_run=dry_run, project=project)


@mcp.tool(annotations=WRITE)
def assign_resources(
    assignments: list[dict[str, Any]],
    create_missing: bool = True,
    project: str | None = None,
) -> dict[str, Any]:
    """Assign resources to tasks.

    Each item: {"task": 12, "resource": "Alice", "units": 0.5}
    units is a fraction of capacity (1.0 = 100%). Missing resources are created
    when create_missing is true.
    """
    return client.assign_resources(assignments, create_missing=create_missing, project=project)


@mcp.tool(annotations=DESTRUCTIVE)
def unassign_resources(assignments: list[dict[str, Any]], project: str | None = None) -> dict[str, Any]:
    """Remove resource assignments. Each item: {"task": 12, "resource": "Alice"}."""
    return client.unassign_resources(assignments, project=project)


@mcp.tool(annotations=READ)
def get_assignments(
    task: int | None = None,
    resource: str | None = None,
    limit: int | None = None,
    project: str | None = None,
) -> dict[str, Any]:
    """List assignments with work, cost, dates, and allocation units, optionally filtered by task or resource."""
    return client.get_assignments(task=task, resource=resource, limit=limit, project=project)


# ---------------------------------------------------------------------------
# Calendars
# ---------------------------------------------------------------------------


@mcp.tool(annotations=READ)
def get_calendars(project: str | None = None) -> dict[str, Any]:
    """List base calendars in the project and identify the project calendar."""
    return client.get_calendars(project=project)


@mcp.tool(annotations=WRITE)
def create_calendar(name: str, copy_from: str = "Standard", project: str | None = None) -> dict[str, Any]:
    """Create a base calendar, copying working times from an existing calendar."""
    return client.create_calendar(name, copy_from=copy_from, project=project)


@mcp.tool(annotations=WRITE)
def add_calendar_exception(
    calendar: str,
    name: str,
    start: str,
    finish: str | None = None,
    working: bool = False,
    project: str | None = None,
) -> dict[str, Any]:
    """Add a holiday or exception date range to a base calendar, e.g. name="Eid", start="2026-03-20"."""
    return client.add_calendar_exception(
        calendar, name, start, finish=finish, working=working, project=project
    )


@mcp.tool(annotations=READ)
def list_calendar_exceptions(calendar: str, project: str | None = None) -> dict[str, Any]:
    """List the exception/holiday entries on a calendar, with their index for deletion."""
    return client.list_calendar_exceptions(calendar, project=project)


@mcp.tool(annotations=DESTRUCTIVE)
def delete_calendar_exception(calendar: str, index: int, project: str | None = None) -> dict[str, Any]:
    """Delete one calendar exception by its index (see list_calendar_exceptions)."""
    return client.delete_calendar_exception(calendar, index, project=project)


@mcp.tool(annotations=WRITE)
def set_working_hours(
    calendar: str,
    day: str,
    shifts: list[list[str]],
    project: str | None = None,
) -> dict[str, Any]:
    """Set weekly working hours for one day of a calendar.

    day: sunday..saturday or 1..7 (1=Sunday). shifts: [["08:00","12:00"],["13:00","17:00"]].
    An empty list makes the day non-working.
    """
    return client.set_working_hours(calendar, day, shifts, project=project)


# ---------------------------------------------------------------------------
# Scheduling, baselines, analysis
# ---------------------------------------------------------------------------


@mcp.tool(annotations=WRITE)
def calculate_project(project: str | None = None) -> dict[str, Any]:
    """Recalculate the schedule and return the new project finish date."""
    return client.calculate(project=project)


@mcp.tool(annotations=DESTRUCTIVE)
def level_resources(project: str | None = None) -> dict[str, Any]:
    """Run Microsoft Project's automatic resource leveling over the whole project."""
    return client.level_resources(project=project)


@mcp.tool(annotations=WRITE)
def save_baseline(number: int = 0, project: str | None = None) -> dict[str, Any]:
    """Save a baseline snapshot (0 = Baseline, 1-10 = Baseline1..10) for variance and earned-value tracking."""
    return client.save_baseline(number=number, project=project)


@mcp.tool(annotations=DESTRUCTIVE)
def clear_baseline(number: int = 0, project: str | None = None) -> dict[str, Any]:
    """Clear a saved baseline (0-10)."""
    return client.clear_baseline(number=number, project=project)


@mcp.tool(annotations=READ)
def get_baseline_variance(number: int = 0, project: str | None = None) -> dict[str, Any]:
    """Compare current task dates, duration, work, and cost against a baseline; reports finish slip per task."""
    return client.baseline_variance(number=number, project=project)


@mcp.tool(annotations=READ)
def get_earned_value(project: str | None = None) -> dict[str, Any]:
    """Per-task BCWS/BCWP/ACWP/SV/CV/SPI/CPI plus project totals. Requires a baseline and progress."""
    return client.earned_value(project=project)


@mcp.tool(annotations=READ)
def get_critical_path(project: str | None = None) -> dict[str, Any]:
    """List non-summary critical tasks in chronological order with slack and predecessors."""
    return client.critical_path(project=project)


@mcp.tool(annotations=READ)
def check_schedule(project: str | None = None) -> dict[str, Any]:
    """Audit the schedule: overdue tasks, negative slack, hard constraints, tasks without resources, tasks without predecessors."""
    return client.check_schedule(project=project)


@mcp.tool(annotations=WRITE)
def simulate_delay(unique_id: int, days: float, project: str | None = None) -> dict[str, Any]:
    """What-if: delay one task by N calendar days, report the impact on project finish and the critical path, then revert.

    The temporary Start No Earlier Than constraint is always removed, even on error.
    """
    return client.simulate_delay(unique_id, days, project=project)


@mcp.tool(annotations=READ)
def get_timephased_data(
    unique_id: int,
    start: str,
    finish: str,
    data_type: str = "work",
    timescale: str = "weekly",
    project: str | None = None,
) -> dict[str, Any]:
    """Period-by-period timephased data for a task.

    data_type: work, cost, actual_work, actual_cost, remaining_work, baseline_work, baseline_cost.
    timescale: daily, weekly, monthly. Keep the date range modest on large projects.
    """
    return client.timephased(
        unique_id,
        start,
        finish,
        data_type=data_type,
        timescale=timescale,
        project=project,
    )


# ---------------------------------------------------------------------------
# Import / export
# ---------------------------------------------------------------------------


@mcp.tool(annotations=WRITE)
def export_csv(path: str, fields: list[str] | None = None, project: str | None = None) -> dict[str, Any]:
    """Write the task table to a UTF-8 CSV at `path` (Excel-friendly BOM included)."""
    return client.export_csv(path, fields=fields, project=project)


@mcp.tool(annotations=WRITE)
def export_pdf(path: str, project: str | None = None) -> dict[str, Any]:
    """Export the project to PDF using Project's own exporter."""
    return client.export_pdf(path, project=project)


@mcp.tool(annotations=WRITE)
def import_project_xml(path: str) -> dict[str, Any]:
    """Open an MS Project XML file as a project in the application."""
    return client.import_xml(path)


@mcp.tool(annotations=READ)
def snapshot_project(
    path: str | None = None, limit: int | None = None, project: str | None = None
) -> dict[str, Any]:
    """Return a JSON snapshot of the project (tasks, resources, metadata), or write it to `path` when given."""
    return client.snapshot(path=path, limit=limit, project=project)


# ---------------------------------------------------------------------------
# Escape hatches
# ---------------------------------------------------------------------------


@mcp.tool(annotations=DESTRUCTIVE)
def run_macro(name: str, args: list[Any] | None = None) -> dict[str, Any]:
    """Run a VBA macro that already exists in the project or global template, e.g. "Module1.BuildWBS"."""
    return client.run_macro(name, args=args)


@mcp.tool(annotations=DESTRUCTIVE)
def run_vba(code: str) -> dict[str, Any]:
    """Execute arbitrary VBA in the project (opt-in). Assign the result to MspMcpEval.

    Example: Dim t: Set t = ActiveProject.Tasks(3): MspMcpEval = t.Finish
    Requires MSPROJECT_MCP_ALLOW_VBA=1 and Project's VBA project trust setting.
    """
    return client.run_vba(code)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------


def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(prog="msproject-mcp")
    parser.add_argument(
        "--transport",
        choices=["stdio", "streamable-http", "sse"],
        default=os.environ.get("MSPROJECT_MCP_TRANSPORT", "stdio"),
        help="stdio (default) for local agents; streamable-http for remote ones",
    )
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8765)
    args = parser.parse_args(argv)

    logging.basicConfig(
        level=os.environ.get("MSPROJECT_MCP_LOG_LEVEL", "WARNING").upper(),
        stream=sys.stderr,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    )

    if args.transport == "stdio":
        mcp.run(transport="stdio")
    else:
        mcp.run(transport=args.transport, host=args.host, port=args.port)


if __name__ == "__main__":
    main()
