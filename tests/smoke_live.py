"""Live integration test. Requires Microsoft Project on Windows.

Creates a throwaway project, exercises the tool surface against the real
application, then closes what it created. Existing open files are left alone.

Run: python tests/smoke_live.py
"""

from __future__ import annotations

import json
import os
import sys
import tempfile
import traceback

from msproject_mcp.core import MSProjectClient

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
except Exception:
    pass

OUT_DIR = tempfile.mkdtemp(prefix="msproject-mcp-test-")
FAILURES: list[str] = []
LAST_CREATED_PROJECT = {"name": None}


def show(label, value, limit=1600):
    print(f"\n=== {label} ===")
    text = json.dumps(value, indent=2, default=str)
    print(text if len(text) <= limit else text[:limit] + "\n... (truncated)")


def step(label, fn):
    print(f"\n--- {label} ---")
    try:
        result = fn()
        show(label, result)
        print("OK")
        return result
    except Exception as exc:
        FAILURES.append(f"{label}: {exc}")
        print(f"FAILED: {exc}")
        traceback.print_exc()
        return None


def expect_error(label, fn, contains=None):
    print(f"\n--- {label} ---")
    try:
        result = fn()
    except Exception as exc:
        text = str(exc)
        if contains and contains not in text:
            FAILURES.append(f"{label}: unexpected error text: {text}")
            print("FAILED (wrong error):", text)
        else:
            print("OK (rejected):", text[:220])
        return
    FAILURES.append(f"{label}: expected an error but the call succeeded")
    print("FAILED: expected an error, got", result)


def main() -> int:
    client = MSProjectClient()
    before = client.status()
    show("status before", before)
    was_running = bool(before.get("running"))

    try:
        return run_steps(client)
    finally:
        if LAST_CREATED_PROJECT["name"]:
            step("close_project", lambda: client.close_project(save=False))
        if not was_running:
            step("quit_msproject", lambda: client.quit_application(save=False))


def run_steps(client: MSProjectClient) -> int:
    project = step(
        "create_project",
        lambda: client.create_project("MCP Smoke Test", start="2026-01-05"),
    )
    if project is None:
        print("\nCannot continue without a project.")
        return 1
    LAST_CREATED_PROJECT["name"] = project.get("name")

    step(
        "update_project",
        lambda: client.update_project({"title": "MCP Smoke Test", "manager": "QA"}),
    )

    created = step(
        "create_tasks",
        lambda: client.create_tasks(
            [
                {"name": "Plan", "duration": "5d"},
                {"name": "Requirements", "duration": "3d", "indent": 1},
                {"name": "Design", "duration": "4d", "indent": 1},
                {"name": "Build", "duration": "10d"},
                {"name": "Test", "duration": "5d"},
            ]
        ),
    )
    if not created:
        return 1
    ids = [row["unique_id"] for row in created["items"]]
    print("task ids:", ids)

    step("get_tasks", lambda: client.get_tasks(limit=50))
    step(
        "get_tasks rich fields",
        lambda: client.get_tasks(
            fields=["name", "duration", "start", "finish", "work", "text1"],
            limit=50,
        ),
    )

    step(
        "set_task_links requirements->design",
        lambda: client.set_task_links(
            ids[2], [{"predecessor": ids[1], "type": "FS", "lag": "1d"}]
        ),
    )
    step(
        "set_task_links design->build",
        lambda: client.set_task_links(ids[3], [{"predecessor": ids[2], "type": "FS"}]),
    )
    step(
        "set_task_links build->test",
        lambda: client.set_task_links(ids[4], [{"predecessor": ids[3], "type": "FS"}]),
    )
    step("get_task_links", lambda: client.get_task_links(ids[3]))

    step(
        "create_resources",
        lambda: client.create_resources(
            [
                {"name": "Alice", "max_units": 1.0, "standard_rate": 85.0, "group": "Eng"},
                {"name": "Bob", "max_units": 0.5, "standard_rate": 70.0, "group": "Eng"},
            ]
        ),
    )
    step("get_resources", lambda: client.get_resources())

    step(
        "assign_resources",
        lambda: client.assign_resources(
            [
                {"task": ids[1], "resource": "Alice", "units": 0.5},
                {"task": ids[2], "resource": "Alice", "units": 1.0},
                {"task": ids[3], "resource": "Bob", "units": 0.5},
            ]
        ),
    )
    step("get_assignments", lambda: client.get_assignments())
    step("get_assignments for task", lambda: client.get_assignments(task=ids[3]))
    step(
        "update_resources",
        lambda: client.update_resources(
            [{"resource": "Alice", "set": {"max_units": 0.75, "standard_rate": 90.0}}]
        ),
    )
    step(
        "unassign_resources",
        lambda: client.unassign_resources([{"task": ids[1], "resource": "Alice"}]),
    )
    step(
        "reassign_resources",
        lambda: client.assign_resources(
            [{"task": ids[1], "resource": "Alice", "units": 0.25}]
        ),
    )

    step("calculate", lambda: client.calculate())
    step("get_critical_path", lambda: client.critical_path())
    step("check_schedule", lambda: client.check_schedule())
    step("level_resources", lambda: client.level_resources())

    step("save_baseline", lambda: client.save_baseline(0))
    step("clear_baseline", lambda: client.clear_baseline(0))
    step("save_baseline again", lambda: client.save_baseline(0))
    step(
        "update_tasks dry_run",
        lambda: client.update_tasks(
            [
                {"unique_id": ids[3], "set": {"duration": "12d", "percent_complete": 25}},
                {"unique_id": ids[4], "set": {"percent_complete": 10}},
            ],
            dry_run=True,
        ),
    )
    step(
        "update_tasks",
        lambda: client.update_tasks(
            [
                {"unique_id": ids[3], "set": {"duration": "12d", "percent_complete": 25}},
                {"unique_id": ids[4], "set": {"percent_complete": 10}},
            ]
        ),
    )
    step("baseline_variance", lambda: client.baseline_variance(0))
    step("earned_value", lambda: client.earned_value())
    step("simulate_delay", lambda: client.simulate_delay(ids[3], 3))

    step(
        "get_timephased_data",
        lambda: client.timephased(ids[3], "2026-01-05", "2026-02-15", "work", "weekly"),
    )

    step(
        "create_calendar",
        lambda: client.create_calendar("MCP Calendar", copy_from="Standard"),
    )
    step(
        "add_calendar_exception",
        lambda: client.add_calendar_exception(
            "MCP Calendar", "Test Holiday", "2026-03-20"
        ),
    )
    step(
        "list_calendar_exceptions",
        lambda: client.list_calendar_exceptions("MCP Calendar"),
    )
    step(
        "delete_calendar_exception",
        lambda: client.delete_calendar_exception("MCP Calendar", 1),
    )
    step(
        "set_working_hours",
        lambda: client.set_working_hours(
            "MCP Calendar", "friday", [["08:00", "12:00"], ["13:00", "16:00"]]
        ),
    )
    step("get_calendars", lambda: client.get_calendars())

    extra = step(
        "create task for deletion",
        lambda: client.create_tasks([{"name": "Temp", "duration": "1d"}]),
    )
    if extra:
        temp_id = extra["items"][0]["unique_id"]
        step("delete_tasks dry_run", lambda: client.delete_tasks([temp_id], dry_run=True))
        step("delete_tasks", lambda: client.delete_tasks([temp_id]))

    step("delete_resources dry_run", lambda: client.delete_resources(["Bob"], dry_run=True))
    step("delete_resources", lambda: client.delete_resources(["Bob"]))
    step("move_task", lambda: client.move_task(ids[3], ids[4]))

    csv_path = os.path.join(OUT_DIR, "tasks.csv")
    mpp_path = os.path.join(OUT_DIR, "smoke.mpp")
    json_path = os.path.join(OUT_DIR, "snapshot.json")
    pdf_path = os.path.join(OUT_DIR, "project.pdf")

    step("save_project_as mpp", lambda: client.save_project_as(mpp_path))
    step("export_csv", lambda: client.export_csv(csv_path))
    step("snapshot", lambda: client.snapshot(json_path))
    step("export_pdf", lambda: client.export_pdf(pdf_path))

    if os.path.exists(csv_path):
        print("\ncsv head:")
        with open(csv_path, encoding="utf-8-sig") as handle:
            for _ in range(3):
                print(handle.readline().rstrip())
    for path in (mpp_path, json_path, pdf_path):
        print(f"{path}: {'exists' if os.path.exists(path) else 'MISSING'} "
              f"({os.path.getsize(path) if os.path.exists(path) else 0} bytes)")

    expect_error(
        "run_vba when disabled",
        lambda: client.run_vba("MspMcpEval = 1"),
        contains="VBA execution is disabled",
    )
    step("get_field_catalog", lambda: client.field_catalog("task"))

    # Cleanup happens in main()'s finally block.
    print("\n================ SUMMARY ================")
    if FAILURES:
        print(f"{len(FAILURES)} failure(s):")
        for failure in FAILURES:
            print(" -", failure)
        return 1
    print("All steps passed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
