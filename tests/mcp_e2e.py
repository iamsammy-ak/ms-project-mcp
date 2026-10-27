"""End-to-end MCP test: spawns the server over stdio and calls real tools.

Requires Microsoft Project. Creates and closes its own throwaway project.

Run: python tests/mcp_e2e.py
"""

from __future__ import annotations

import asyncio
import json
import os
import sys

from mcp import ClientSession
from mcp.client.stdio import StdioServerParameters, stdio_client

FAILURES: list[str] = []


def think(label, result):
    if getattr(result, "is_error", False):
        text = " ".join(
            getattr(block, "text", "") for block in getattr(result, "content", [])
        )
        print(f"  FAILED {label}: {text[:300]}")
        FAILURES.append(f"{label}: {text[:300]}")
        return None
    payload = getattr(result, "structured_content", None)
    if payload is None:
        text = " ".join(
            getattr(block, "text", "") for block in getattr(result, "content", [])
        )
        try:
            payload = json.loads(text)
        except Exception:
            payload = {"text": text}
    return payload


def summarise(payload, keys):
    if not isinstance(payload, dict):
        return payload
    return {key: payload.get(key) for key in keys if key in payload}


async def main() -> int:
    params = StdioServerParameters(
        command=sys.executable,
        args=["-m", "msproject_mcp"],
        env={**os.environ, "MSPROJECT_MCP_LOG_LEVEL": "WARNING"},
    )

    async with stdio_client(params) as (read, write):
        async with ClientSession(read, write) as session:
            await session.initialize()

            tools = await session.list_tools()
            names = [tool.name for tool in tools.tools]
            print(f"tools: {len(names)}")
            for required in ("get_tasks", "create_tasks", "assign_resources", "export_pdf"):
                if required not in names:
                    FAILURES.append(f"missing tool {required}")
            schema = next(t for t in tools.tools if t.name == "create_tasks")
            print("create_tasks schema keys:", list(schema.input_schema.get("properties", {})))

            status = think("status", await session.call_tool("msproject_status", {}))
            print("status:", summarise(status, ["running", "open_project_count"]))

            project = think(
                "create_project",
                await session.call_tool(
                    "create_project", {"title": "MCP E2E", "start": "2026-02-02"}
                ),
            )
            print("project:", summarise(project, ["name", "title", "start"]))
            if project is None:
                return 1

            created = think(
                "create_tasks",
                await session.call_tool(
                    "create_tasks",
                    {
                        "tasks": [
                            {"name": "Phase 1", "duration": "5d"},
                            {"name": "Draft", "duration": "3d", "indent": 1},
                            {"name": "Review", "duration": "2d", "indent": 1},
                            {"name": "Ship", "duration": "1d"},
                        ]
                    },
                ),
            )
            ids = [row["unique_id"] for row in created["items"]]
            print("created ids:", ids)

            linked = think(
                "set_task_links",
                await session.call_tool(
                    "set_task_links",
                    {"unique_id": ids[3], "links": [{"predecessor": ids[2], "type": "FS"}]},
                ),
            )
            print("predecessors:", linked.get("predecessors") if linked else None)

            think(
                "create_resources",
                await session.call_tool(
                    "create_resources", {"resources": [{"name": "Alice", "max_units": 1.0}]}
                ),
            )
            think(
                "assign_resources",
                await session.call_tool(
                    "assign_resources",
                    {"assignments": [{"task": ids[1], "resource": "Alice", "units": 0.5}]},
                ),
            )
            assignments = think(
                "get_assignments", await session.call_tool("get_assignments", {})
            )
            print("assignments:", assignments.get("returned") if assignments else None)

            tasks = think(
                "get_tasks",
                await session.call_tool(
                    "get_tasks", {"fields": ["name", "duration", "start", "finish"], "limit": 10}
                ),
            )
            print("tasks returned:", tasks.get("returned"))

            think("calculate", await session.call_tool("calculate_project", {}))
            detailed = think(
                "get_tasks detail",
                await session.call_tool(
                    "get_tasks",
                    {
                        "fields": ["name", "critical", "total_slack", "start", "finish"],
                        "limit": 10,
                    },
                ),
            )
            print(
                "critical flags:",
                [
                    (row["name"], row.get("critical"), row.get("total_slack"))
                    for row in detailed["items"]
                ],
            )

            critical = think(
                "get_critical_path", await session.call_tool("get_critical_path", {})
            )
            print("critical tasks:", critical.get("critical_tasks") if critical else None)

            audit = think("check_schedule", await session.call_tool("check_schedule", {}))
            print("issues:", audit.get("issue_count") if audit else None)

            print("--- expected error path ---")
            bad = await session.call_tool("get_task", {"unique_id": 999999})
            bad_text = " ".join(
                getattr(block, "text", "") for block in getattr(bad, "content", [])
            )
            print("is_error:", bad.is_error)
            print("error text:", bad_text[:220])
            if not bad.is_error:
                FAILURES.append("get_task with a bogus id should fail")
            if "No task with id" not in bad_text:
                FAILURES.append(f"error text should name the missing task, got: {bad_text[:160]}")

            print("--- cleanup ---")
            think("close_project", await session.call_tool("close_project", {"save": False}))
            think("quit", await session.call_tool("quit_msproject", {"save": False}))

    print("\n================ SUMMARY ================")
    if FAILURES:
        for failure in FAILURES:
            print(" -", failure)
        return 1
    print("End-to-end MCP test passed.")
    return 0


if __name__ == "__main__":
    sys.exit(asyncio.run(main()))
