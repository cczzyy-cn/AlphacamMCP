"""
MCP bridge error types — structured errors with codes and recoverability hints.
"""


class ToolError(Exception):
    """Structured error from an MCP tool handler."""

    def __init__(self, message: str, code: str = "TOOL_ERROR",
                 recoverable: bool = False):
        self.code = code
        self.recoverable = recoverable
        super().__init__(message)

    def to_dict(self) -> dict:
        return {
            "error": self.args[0],
            "code": self.code,
            "recoverable": self.recoverable,
        }


class AlphaCAMNotConnected(ToolError):
    """AlphaCAM is not available."""

    def __init__(self):
        super().__init__(
            "AlphaCAM is not connected or not running. "
            "Ensure AlphaCAM is installed and try again.",
            code="NOT_CONNECTED",
            recoverable=True,
        )


class ToolArgumentError(ToolError):
    """Invalid or missing tool argument."""

    def __init__(self, message: str):
        super().__init__(message, code="INVALID_ARGS", recoverable=True)


class ToolComError(ToolError):
    """COM automation call failed, possibly transient."""

    def __init__(self, message: str, retries: int = 0):
        super().__init__(
            f"COM error{' after {retries} retries' if retries else ''}: {message}",
            code="COM_ERROR",
            recoverable=True,
        )


class ToolNotFoundError(ToolError):
    """Unknown tool name."""

    def __init__(self, name: str):
        super().__init__(f"Unknown tool: {name}", code="NOT_FOUND")
