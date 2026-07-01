"""
AlphaCAM MCP Server — entry point.

Usage:
    python server.py                 # stdio transport
    python server.py --port 8080     # SSE transport
    python server.py --progid am5axaps.Application
"""

from __future__ import annotations

import argparse
import logging

from mcp.server import Server, NotificationOptions
from mcp.server.models import InitializationOptions
from mcp.types import CallToolResult
import mcp.server.stdio

from mcp_bridge import TOOLS, handle_tool, set_prog_id
from mcp_bridge.config import get_acPort

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
log = logging.getLogger("alphacam-bridge")


def create_server() -> Server:
    server = Server("alphacam-bridge")

    @server.list_tools()
    async def list_tools():
        return TOOLS

    @server.call_tool()
    async def call_tool(name: str, arguments: dict | None) -> CallToolResult:
        return await handle_tool(name, arguments)

    return server


async def run_stdio():
    server = create_server()
    async with mcp.server.stdio.stdio_server() as (read, write):
        await server.run(
            read, write,
            InitializationOptions(
                server_name="alphacam-bridge",
                server_version="1.1.0",
                capabilities=server.get_capabilities(
                    notification_options=NotificationOptions(),
                    experimental_capabilities={},
                ),
            ),
        )


def main():
    parser = argparse.ArgumentParser(
        description="AlphaCAM MCP Bridge Server")
    parser.add_argument("--port", type=int, default=None,
                        help="Port for SSE transport (omit for stdio)")
    parser.add_argument("--progid", type=str, default=None,
                        help="COM ProgID override (e.g. am5axaps.Application)")
    args = parser.parse_args()

    if args.progid:
        set_prog_id(args.progid)

    port = args.port or get_acPort()

    if port:
        # SSE mode
        from starlette.applications import Starlette
        from starlette.routing import Mount, Route
        from mcp.server.sse import SseServerTransport
        import uvicorn

        server = create_server()
        sse = SseServerTransport("/messages/")

        async def handle_sse(request):
            async with sse.connect_sse(
                request.scope, request.receive, request._send,
            ) as (read, write):
                await server.run(
                    read, write,
                    InitializationOptions(
                        server_name="alphacam-bridge",
                        server_version="1.1.0",
                        capabilities=server.get_capabilities(
                            notification_options=NotificationOptions(),
                            experimental_capabilities={},
                        ),
                    ),
                )

        app = Starlette(
            routes=[
                Route("/sse", endpoint=handle_sse),
                Mount("/messages/", app=sse.handle_post_message),
            ],
        )
        log.info(f"Starting SSE server on port {port}")
        uvicorn.run(app, host="127.0.0.1", port=port)
    else:
        # stdio mode
        import asyncio
        log.info("Starting stdio server")
        asyncio.run(run_stdio())


if __name__ == "__main__":
    main()
