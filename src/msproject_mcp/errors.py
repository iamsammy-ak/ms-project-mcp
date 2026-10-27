"""Exceptions surfaced to MCP clients.

Subclassing the SDK's ToolError means the message reaches the model as a tool
error it can read and react to, rather than a generic "unexpected exception"
with a server-side traceback.
"""

from __future__ import annotations

from mcp.server.mcpserver.exceptions import ToolError


class MSProjectMCPError(ToolError):
    """Base class for errors that should reach the model verbatim."""


class ProjectNotRunning(MSProjectMCPError):
    pass


class NoProjectOpen(MSProjectMCPError):
    pass


class NotFound(MSProjectMCPError):
    pass


class InvalidInput(MSProjectMCPError):
    pass


class FieldNotWritable(InvalidInput):
    pass


class ComTimeout(MSProjectMCPError):
    pass


class VbaDisabled(MSProjectMCPError):
    pass
