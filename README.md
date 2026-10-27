<div align="center">

# ms-project-mcp

**Drive Microsoft Project desktop from any MCP-capable agent.**

Create the plan, wire the dependencies, level the resources, save baselines, audit the critical path, export the report — all against the real `.mpp` open in Project.

[![Python](https://img.shields.io/badge/python-3.10%2B-3776AB?logo=python&logoColor=white)](https://www.python.org/)
[![Platform](https://img.shields.io/badge/platform-Windows-0078D4?logo=windows&logoColor=white)](#requirements)
[![MCP](https://img.shields.io/badge/MCP-2.x-6B4FBB)](https://modelcontextprotocol.io/)
[![License](https://img.shields.io/badge/license-MIT-green)](LICENSE)
[![Tools](https://img.shields.io/badge/tools-51-brightgreen)](#tool-reference)

</div>

---

## Why this exists

Microsoft Project desktop has no public HTTP API. The only way to automate it is
COM — a single-threaded, apartment-bound automation surface that is easy to get
subtly wrong. This server turns that surface into **51 typed MCP tools** with a
stable id scheme, field-name-based reads and writes, dry runs, and real error
messages, so an agent can do project management work instead of fighting COM.

It talks to the **application you have open**, not a cloud copy. If you can click
it, the agent can usually do it — and you can watch it happen.

---

## What it can do

| Area | Capabilities |
| --- | --- |
| **Projects** | Attach to the running window, list/switch open files, create, open `.mpp`/`.xml`/`.mpt`, save, save-as, close, quit, show/hide |
| **Tasks** | Bulk create with WBS nesting, update, delete, indent/outdent, reorder, read/write almost every field, custom fields, paging and filters |
| **Dependencies** | FS / SS / FF / SF links with lag or lead, append or replace, parsed successors |
| **Resources** | Work, material, and cost resources; rates, max units, calendars, groups; assign with allocation units, reassign, unassign |
| **Calendars** | Base calendars, holidays and exception ranges, per-day working shifts, project calendar |
| **Baselines & value** | Save/clear baselines 0–10, per-task variance, earned value (BCWS/BCWP/ACWP/SPI/CPI) |
| **Analysis** | Critical path, schedule audit, what-if delay with automatic revert, timephased data |
| **Export** | CSV, PDF, full JSON snapshot, open existing Project XML |
| **Escape hatches** | Run an existing VBA macro, or inject arbitrary VBA (opt-in) with results and errors returned |

---

## Architecture

```
┌──────────────────────────┐
│  Agent (Command Code,    │
│  Claude, Cursor, …)      │
└────────────┬─────────────┘
             │  stdio · JSON-RPC (MCP)
┌────────────▼─────────────┐
│  MCP tools               │  thin, typed, annotated
├──────────────────────────┤
│  MSProjectClient         │  domain logic, field registry
├──────────────────────────┤
│  ComBridge               │  ONE STA thread · serialized queue
│                          │  busy-retry · timeouts
├──────────────────────────┤
│  WINPROJ.EXE             │  attaches, or launches if needed
└──────────────────────────┘
```

Office COM objects are apartment-threaded. Calling them from several threads
produces intermittent `RPC_E_WRONG_THREAD` and *"Call was rejected by callee"*
errors, so **every** COM call is marshalled onto one dedicated thread and
executed serially. If you click around in Project while an agent works, rejected
calls are retried instead of failing.

---

## Requirements

- **Windows** with desktop Microsoft Project installed (2016 / 2019 / 2021 / Microsoft 365).
  Verified end to end against **Project 16.0 (Microsoft 365)**.
- **Python 3.10+** (developed on 3.13).
- No project file needs to be open at startup.

> Project **Online / for the Web** is a different product with a different API
> (Graph/CSOM). This server targets the desktop application only.

---

## Install

```powershell
git clone https://github.com/iamsammy-ak/ms-project-mcp.git
cd ms-project-mcp
python -m venv .venv
.\.venv\Scripts\python -m pip install -e .
```

Check that it can see Project:

```powershell
.\.venv\Scripts\python -c "from msproject_mcp.core import MSProjectClient; print(MSProjectClient().status())"
```

Prefer a global command? `pipx install .` gives you `msproject-mcp` on `PATH`.
To ship it to another machine, build a wheel and copy that:

```powershell
python -m pip wheel --no-deps -w dist .
# elsewhere: pip install msproject_mcp-0.1.0-py3-none-any.whl
```

---

## Connect your agent

Use the absolute path to your interpreter (or the `msproject-mcp` shim).

### Command Code

```powershell
cmdc mcp add --scope user --env MSPROJECT_MCP_ALLOW_VBA=1 msproject -- "C:\path\to\ms-project-mcp\.venv\Scripts\python.exe" -m msproject_mcp
```

### Claude Desktop / Claude Code / Cursor / VS Code

```json
{
  "mcpServers": {
    "msproject": {
      "command": "C:\\path\\to\\ms-project-mcp\\.venv\\Scripts\\python.exe",
      "args": ["-m", "msproject_mcp"],
      "env": { "MSPROJECT_MCP_ALLOW_VBA": "1" }
    }
  }
}
```

Remote agents can use the HTTP transport instead:

```powershell
python -m msproject_mcp --transport streamable-http --port 8765
```

Keep it bound to `127.0.0.1` unless you understand the exposure — it drives your
desktop applications.

---

## Try it

With Project open (or not — the agent can start it):

```text
"What's the critical path in the open project?"

"Create a launch plan: Discovery 5d, Design 8d, Build 15d, Test 6d,
 UAT 4d, Launch 1d — link them finish-to-start, indent Build..Test
 under a Delivery phase."

"Alice is 50% available and Bob is out for two weeks. Fix the
 resource overload and tell me what slipped."

"Save baseline 0, then show every task that is more than 3 days late."

"Export the schedule to C:\Temp\launch.csv and a PDF next to it."

"Add our Eid holidays to the Standard calendar for 2026."
```

---

## Tool reference

<details open>
<summary><b>Session and project lifecycle</b></summary>

| Tool | What it does |
| --- | --- |
| `msproject_status` | Is Project running, which files are open, which is active, is VBA enabled |
| `list_open_projects` | Every open project with task/resource counts |
| `use_project` | Switch active project by index, name, fragment, or path |
| `create_project` | New blank project with title and start date |
| `open_project` | Open `.mpp`, `.xml`, or `.mpt` |
| `save_project` | Save in place |
| `save_project_as` | Save as `.mpp` or `.mpt` |
| `close_project` | Close a project, optionally saving |
| `quit_msproject` | Quit the application |
| `set_project_window_visible` | Show or hide the Project window |

</details>

<details>
<summary><b>Project properties</b></summary>

| Tool | What it does |
| --- | --- |
| `get_project` | Title, dates, calendar, minutes/day and any requested fields |
| `update_project` | Update title, manager, status date, calendar, scheduling options, currency… |
| `get_field_catalog` | Every valid field name for tasks, resources, or the project |

</details>

<details>
<summary><b>Tasks</b></summary>

| Tool | What it does |
| --- | --- |
| `get_tasks` | List with name/level/summary/milestone filters and paging |
| `get_task` | Read one task with any field set |
| `create_tasks` | Append tasks in order; `indent` builds the WBS; auto-scheduled by default |
| `update_tasks` | Bulk field edits with `dry_run` preview and before/after values |
| `delete_tasks` | Delete by id, with preview |
| `indent_task` | Promote/demote in the WBS |
| `move_task` | Reorder a row (uses Project's cut/paste) |

</details>

<details>
<summary><b>Dependencies</b></summary>

| Tool | What it does |
| --- | --- |
| `set_task_links` | FS/SS/FF/SF links with lag/lead; replace or append |
| `get_task_links` | Parsed predecessors and successors |

</details>

<details>
<summary><b>Resources and assignments</b></summary>

| Tool | What it does |
| --- | --- |
| `get_resources` | Resource pool with filters and paging |
| `create_resources` | Add work/material/cost resources |
| `update_resources` | Change rates, units, calendar, grouping, custom fields |
| `delete_resources` | Remove resources and their assignments |
| `assign_resources` | Assign with units; creates missing resources on request |
| `unassign_resources` | Remove assignments |
| `get_assignments` | Work, actual/remaining work, cost, dates, % complete |

</details>

<details>
<summary><b>Calendars</b></summary>

| Tool | What it does |
| --- | --- |
| `get_calendars` | Base calendars and the project calendar |
| `create_calendar` | New calendar copied from an existing one |
| `add_calendar_exception` | Holiday or exception date range |
| `list_calendar_exceptions` | Exceptions with their index |
| `delete_calendar_exception` | Remove an exception |
| `set_working_hours` | Per-day shifts, or mark a day non-working |

</details>

<details>
<summary><b>Scheduling and analysis</b></summary>

| Tool | What it does |
| --- | --- |
| `calculate_project` | Recalculate and return the new finish date |
| `level_resources` | Run Project's automatic leveling |
| `save_baseline` / `clear_baseline` | Baseline 0–10 |
| `get_baseline_variance` | Finish slip, duration/work/cost variance per task |
| `get_earned_value` | BCWS/BCWP/ACWP/SV/CV/SPI/CPI and totals |
| `get_critical_path` | Critical tasks in chronological order with slack |
| `check_schedule` | Overdue, negative slack, hard constraints, missing resources/links |
| `simulate_delay` | What-if delay, then revert automatically |
| `get_timephased_data` | Per-day/week/month work, cost, actual, remaining, baseline |

</details>

<details>
<summary><b>Import, export, escape hatches</b></summary>

| Tool | What it does |
| --- | --- |
| `export_csv` | UTF-8 CSV of the task table |
| `export_pdf` | PDF through Project's own exporter |
| `snapshot_project` | JSON snapshot of tasks, resources, and metadata |
| `import_project_xml` | Open a Project XML file |
| `run_macro` | Run an existing VBA macro |
| `run_vba` | Inject and run VBA (opt-in) |

</details>

---

## Fields instead of a tool per property

Rather than hundreds of narrow tools, task/resource/project tools accept a
`fields` list keyed by stable names. `get_field_catalog` returns the whole set.

```text
get_tasks(fields=["name", "duration", "start", "finish", "text1", "number1"])
update_tasks([{"unique_id": 12, "set": {"duration": "8d", "percent_complete": 40}}])
```

Conventions:

- **Ids** accept `UniqueID` (stable across edits) or the visible row `ID`.
  UniqueID is matched first; every row returns both.
- **Durations** accept `"5d"`, `"12h"`, `"480m"`, or raw minutes. Reads return
  both `duration: "5d"` and `duration_minutes: 2400`.
- **Percentages** are `0–100`; a non-integer `0.0–1.0` value is a fraction.
- **New tasks are auto-scheduled** (`manual: false`) so links and durations
  recalculate. Pass `"manual": true` to pin a task's dates.
- Every bulk mutation supports `dry_run`.

Custom fields are first-class: `text1–text30`, `number1–number20`,
`flag1–flag20`, `date1–date10`, `cost1–cost10`, `duration1–duration10`.

---

## VBA escape hatch

`run_vba` closes the last gap: dialog-only operations and expressions no COM
property exposes. It is off by default.

1. Start the server with `MSPROJECT_MCP_ALLOW_VBA=1`.
2. In Project: **File → Options → Trust Center → Trust Center Settings →
   Macro Settings → Trust access to the VBA project object model**.

The code is injected into a temporary module, executed, and the module is
removed. Assign your result to `MspMcpEval`:

```vba
Dim t As Task
Set t = ActiveProject.Tasks(3)
MspMcpEval = t.Name & " / " & t.Finish
```

Runtime errors come back as readable tool errors (`VBA error: 11: Division by zero`).

---

## Configuration

| Variable | Default | Purpose |
| --- | --- | --- |
| `MSPROJECT_MCP_VISIBLE` | `1` | Show the Project window while working |
| `MSPROJECT_MCP_ALLOW_VBA` | `0` | Enable `run_vba` |
| `MSPROJECT_MCP_TIMEOUT` | `180` | Seconds before a single COM call is abandoned |
| `MSPROJECT_MCP_PROJECT` | — | Project to use when several are open |
| `MSPROJECT_MCP_MAX_ROWS` | `1000` | Row cap for list tools |
| `MSPROJECT_MCP_TRANSPORT` | `stdio` | `stdio`, `streamable-http`, or `sse` |
| `MSPROJECT_MCP_LOG_LEVEL` | `WARNING` | Logs go to stderr only |

---

## What working with Project 16 actually taught us

These are baked into the code, and they are the difference between a demo and a
tool that works:

| Behaviour | Consequence |
| --- | --- |
| `Tasks.Add` creates **manually scheduled** tasks | Links and durations would never propagate; new tasks default to auto-scheduled |
| `Task.Indent()` does not exist | WBS moves use `OutlineLevel`, falling back to `OutlineIndent` |
| `Project.Assignments` is not exposed | Assignments are created with `task.Assignments.Add(TaskID, ResourceID, Units)` positionally |
| PDF export lives on `Application.DocumentExport(Filename, 0)` | `Project.ExportAsFixedFormat` rejects every argument combination |
| Saving to XML through COM opens a modal dialog | XML export is intentionally not exposed; CSV/PDF/snapshot are |
| Rates come back localized (`"₹ 85.00/hr"`) | Numeric fields parse currency strings and also return the raw text |
| Dates arrive UTC-tagged but represent local wall-clock time | Converting to local (not stripping the tz) keeps dates correct |
| `Application.Run` cannot return a Function value | VBA results are written to a temp file and read back |

---

## Limitations

- **One automation client per Project instance.** A modal dialog in Project
  blocks calls until dismissed; they surface as a clear timeout.
- **No undo.** Project's COM undo is one level deep and unreliable. Use
  `dry_run` and `snapshot_project` before destructive batches.
- **`save_project_as` rebinds the open project**, exactly like Project's own
  Save As. Only `.mpp` and `.mpt` are supported.
- **`move_task` uses cut/paste**, so it touches the Windows clipboard.
- **Dialog-only features** (recurring-task insert, report designer) have no COM
  entry point; `run_vba` covers most of them.
- **Timephased queries** over multi-year ranges are slow — keep ranges modest.
- **Windows only**, and desktop Project only.

---

## Testing

```powershell
.\.venv\Scripts\python tests\smoke_live.py   # 47 steps against real Project
.\.venv\Scripts\python tests\mcp_e2e.py      # real server over stdio MCP
```

`smoke_live.py` launches Project, creates a throwaway schedule, exercises every
tool group, and closes only what it created. `mcp_e2e.py` spawns the server over
stdio and drives it through an MCP client, including the error path.

Both require Microsoft Project and are therefore not run in CI.

---

## Contributing

Issues and PRs are welcome. Please keep the invariants:

- All COM access goes through `ComBridge` — never call COM from another thread.
- New fields go in `fields.py`; new behaviour goes in `core.py`; tools stay thin.
- Add a live test step for anything that touches Project.

---

## License

[MIT](LICENSE) © 2026 iamsammy-ak
