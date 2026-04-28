"""FastMCP server entry point.

Selects the right backend for the current OS, registers all tools, and
runs the MCP stdio server loop.
"""

from __future__ import annotations

import argparse
import asyncio
import logging
import sys

from .backends import select_backend
from .errors import UnsupportedPlatformError

log = logging.getLogger(__name__)


def _build_server() -> tuple[object, object]:
    try:
        from mcp.server.fastmcp import FastMCP  # type: ignore[import-not-found]
    except ImportError as exc:
        raise RuntimeError(
            "mcp package not installed. Run `pip install mcp` or `pip install -e .`."
        ) from exc

    mcp = FastMCP("mega-outlook-mcp")
    backend = select_backend()

    from .tools import register_all

    register_all(mcp, backend)
    return mcp, backend


async def _amain() -> int:
    try:
        mcp, backend = _build_server()
    except UnsupportedPlatformError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2

    try:
        # FastMCP.run() is synchronous and blocking; run it in a thread so we
        # can still `await backend.close()` cleanly on shutdown.
        await asyncio.to_thread(mcp.run)  # type: ignore[attr-defined]
        return 0
    finally:
        try:
            await backend.close()  # type: ignore[attr-defined]
        except Exception:  # noqa: BLE001
            pass


def main() -> None:
    parser = argparse.ArgumentParser(
        prog="mega-outlook-mcp",
        description="Cross-platform MCP server for local Outlook automation (Windows COM + macOS AppleScript).",
    )
    parser.add_argument(
        "--log-level",
        default="INFO",
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
        help="Logging verbosity (default: INFO).",
    )
    args = parser.parse_args()
    logging.basicConfig(
        level=args.log_level, format="%(asctime)s %(levelname)s %(name)s: %(message)s"
    )
    exit_code = asyncio.run(_amain())
    sys.exit(exit_code)


if __name__ == "__main__":
    main()

