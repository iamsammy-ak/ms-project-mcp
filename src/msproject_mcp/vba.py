"""Gated VBA escape hatch.

This is how the server reaches anything the typed tools do not cover. It
requires Project's "Trust access to the VBA project object model" setting and
the MSPROJECT_MCP_ALLOW_VBA flag; both are opt-in.

Design note: Project's ``Application.Run`` executes Subs but cannot return a
Function's value to COM. So we inject one Sub, let the caller assign its result
to a public module variable, and write that value to a temp file which Python
reads back.
"""

from __future__ import annotations

import os
import tempfile

from .errors import MSProjectMCPError

_PROC = "MspMcpRun"
_MODULE = "MspMcpModule"


def _vb_project(app, proj):
    try:
        return proj.VBProject
    except Exception:
        pass
    try:
        return app.VBE.ActiveVBProject
    except Exception:
        pass
    raise MSProjectMCPError(
        "Cannot access the VBA project. Enable 'Trust access to the VBA project "
        "object model' in Project > Options > Trust Center > Macro Settings."
    )


def _wrap(code: str, result_path: str) -> str:
    body = code.replace("\r\n", "\n").replace("\r", "\n")
    indented = "\n".join(f"    {line}" if line.strip() else "" for line in body.split("\n"))
    target = result_path.replace('"', '""')
    return (
        "Public MspMcpEval As Variant\n\n"
        f"Public Sub {_PROC}()\n"
        "    On Error GoTo MspMcpErr\n"
        '    MspMcpEval = ""\n'
        f"{indented}\n"
        f'    Open "{target}" For Output As #1\n'
        "    Print #1, CStr(MspMcpEval)\n"
        "    Close #1\n"
        "    Exit Sub\n"
        "MspMcpErr:\n"
        "    Dim mspErrNum As Long, mspErrDesc As String\n"
        "    mspErrNum = Err.Number\n"
        "    mspErrDesc = Err.Description\n"
        "    On Error Resume Next\n"
        f'    Open "{target}" For Output As #1\n'
        '    Print #1, "ERROR " & mspErrNum & ": " & mspErrDesc\n'
        "    Close #1\n"
        "End Sub\n"
    )


def execute_vba(app, proj, code: str):
    if not str(code or "").strip():
        raise MSProjectMCPError("Provide VBA code to run.")

    vbp = _vb_project(app, proj)
    handle, result_path = tempfile.mkstemp(prefix="msp-vba-", suffix=".txt")
    os.close(handle)
    os.unlink(result_path)

    module = vbp.VBComponents.Add(1)  # vbext_ct_StdModule
    try:
        try:
            module.Name = _MODULE
        except Exception:
            pass
        module_name = module.Name
        module.CodeModule.AddFromString(_wrap(code, result_path))
        app.Run(_PROC)

        if not os.path.exists(result_path):
            raise MSProjectMCPError(
                "The VBA ran but produced no result file. The macro may have been "
                "blocked by Project's macro security settings."
            )
        with open(result_path, "r", encoding="mbcs", errors="replace") as handle_in:
            text = handle_in.read().strip()
        if text.startswith("ERROR "):
            raise MSProjectMCPError(f"VBA error: {text[len('ERROR '):]}")
        return {"module": module_name, "result": text}
    finally:
        try:
            vbp.VBComponents.Remove(module)
        except Exception:
            pass
        try:
            os.unlink(result_path)
        except OSError:
            pass
