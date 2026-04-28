# mega-outlook-mcp

[![PyPI](https://img.shields.io/pypi/v/mega-outlook-mcp)](https://pypi.org/project/mega-outlook-mcp/)
[![Python](https://img.shields.io/pypi/pyversions/mega-outlook-mcp)](https://pypi.org/project/mega-outlook-mcp/)
[![License](https://img.shields.io/pypi/l/mega-outlook-mcp)](LICENSE)

The most-comprehensive cross-platform MCP server for local Microsoft Outlook
automation. **64 tools** across email, threading, calendar, contacts, tasks,
notes, inbox rules, automatic replies, free/busy, the Global Address List,
and Exchange-specific surface — all behind a single Python package that runs
on **both Windows and macOS**.

The server is tooling only — scheduling and the agent loop live in your MCP
client (LM Studio, Claude Desktop, or any MCP-compatible harness).

## Why mega-outlook-mcp?

Two great community MCP servers already exist for Outlook:

- [`Aanerud/outlook-desktop-mcp`](https://github.com/Aanerud/outlook-desktop-mcp) — 29 tools, Windows-only, Node.js
- [`hasan-imam/mcp-outlook-applescript`](https://github.com/hasan-imam/mcp-outlook-applescript) — 49 tools, macOS-only, Node.js

`mega-outlook-mcp` is the union of both, plus more, in one Python package
with a unified Backend Protocol. Highlights:

- **One install for both OSes.** `pip install mega-outlook-mcp` works the
  same on Windows and macOS; the right backend is selected at startup.
- **Agent-friendly composite tools.** `outlook_summarize_inbox`,
  `outlook_extract_action_items`, `outlook_find_unanswered`,
  `outlook_meeting_prep`, `outlook_relationship_graph`, and more — single
  calls instead of N round trips.
- **Exchange-specific surface neither reference repo covers.** Out-of-office,
  signatures, inbox rules, free/busy lookup, room finder, GAL search,
  delegated mailboxes, public folders, mailbox quota.
- **Built-in compatibility self-test.** `outlook_diagnostics` probes every
  Outlook field this server depends on and tells you exactly which tools
  break if Microsoft renames something.

## Tool catalog (64 tools)

<details>
<summary><b>Email — extraction, threading, write, organize</b> (24)</summary>

`outlook_get_emails_in_time_range`, `outlook_get_conversation_thread`,
`outlook_get_email_full_metadata`, `outlook_search_emails`,
`outlook_save_attachment`, `outlook_get_thread_metadata`,
`outlook_send_email`, `outlook_create_draft`, `outlook_reply`,
`outlook_reply_all`, `outlook_forward`, `outlook_mark_email_read`,
`outlook_set_email_flag`, `outlook_set_email_categories`,
`outlook_move_email`, `outlook_archive_email`, `outlook_delete_email`,
`outlook_junk_email`, `outlook_summarize_inbox`,
`outlook_extract_action_items`, `outlook_find_unanswered`,
`outlook_find_promised_actions`, `outlook_threadify`,
`outlook_relationship_graph`
</details>

<details>
<summary><b>Folders</b> (6)</summary>

`outlook_list_folders`, `outlook_create_folder`, `outlook_rename_folder`,
`outlook_move_folder`, `outlook_delete_folder`, `outlook_empty_folder`
</details>

<details>
<summary><b>Calendar</b> (7)</summary>

`outlook_list_calendar_events`, `outlook_get_calendar_event`,
`outlook_create_calendar_event`, `outlook_update_calendar_event`,
`outlook_delete_calendar_event`, `outlook_respond_to_event`,
`outlook_meeting_prep`
</details>

<details>
<summary><b>Contacts, tasks, notes</b> (9)</summary>

`outlook_list_contacts`, `outlook_search_contacts`, `outlook_get_contact`,
`outlook_list_tasks`, `outlook_search_tasks`, `outlook_get_task`,
`outlook_list_notes`, `outlook_search_notes`, `outlook_get_note`
</details>

<details>
<summary><b>Mailbox & accounts</b> (4)</summary>

`outlook_get_mailbox_info`, `outlook_list_accounts`,
`outlook_get_unread_count`, `outlook_get_mailbox_quota`
</details>

<details>
<summary><b>Exchange-specific (Windows-COM-only; Mac returns ERROR-MAC-Support-Unavailable)</b> (10)</summary>

`outlook_get_out_of_office`, `outlook_set_out_of_office`,
`outlook_get_signature`, `outlook_set_signature`, `outlook_list_rules`,
`outlook_toggle_rule`, `outlook_calendar_freebusy`,
`outlook_meeting_room_finder`, `outlook_gal_search`,
`outlook_list_delegated_mailboxes`, `outlook_list_public_folders`
</details>

<details>
<summary><b>Utility</b> (4)</summary>

`outlook_get_current_time`, `outlook_write_file`, `outlook_diagnostics`
</details>

## Install

```bash
pip install mega-outlook-mcp
```

Or with `uv` / `pipx`:

```bash
uv tool install mega-outlook-mcp
# or
pipx install mega-outlook-mcp
```

`pywin32` is installed automatically on Windows via a platform marker; nothing
extra is needed on macOS (`osascript` ships with the OS).

## Wire up an MCP client

Add the server to your MCP client config. Examples:

**Claude Desktop** (`~/Library/Application Support/Claude/claude_desktop_config.json` on macOS, `%APPDATA%\Claude\claude_desktop_config.json` on Windows):

```json
{
  "mcpServers": {
    "outlook": {
      "command": "mega-outlook-mcp"
    }
  }
}
```

**LM Studio**: Settings → MCP → add a server with command `mega-outlook-mcp`.

If the GUI app can't find `mega-outlook-mcp` on PATH (common on macOS), use the absolute path:

```json
{
  "mcpServers": {
    "outlook": {
      "command": "/Users/you/.local/bin/python3",
      "args": ["-m", "mega_outlook_mcp.server"]
    }
  }
}
```

## Platform requirements

| OS | Outlook | Notes |
|---|---|---|
| Windows 10/11 | Outlook Desktop (classic) signed in | `pywin32` installs automatically |
| macOS 13+ | Outlook for Mac classic (16.x) | "New Outlook" not supported — Apple sandboxing blocks the AppleScript surface; use the classic UI toggle |

## First-call sanity check

Tell your agent:

> Call `outlook_diagnostics`, then `outlook_get_mailbox_info`.

`outlook_diagnostics` should return `status: HEALTHY`. If it returns
`DEGRADED:<n>_fields` or `BROKEN`, read `affected_tools` and `notes` for
the cause. Most common: Outlook is closed, or the user is on "New Outlook"
on macOS.

## Future-proofing against Outlook updates

`outlook_diagnostics` probes every Outlook field this server depends on
against a baseline manifest at `src/mega_outlook_mcp/baseline/outlook_baseline.json`.
Run it on a schedule (`launchd`, Task Scheduler, cron) so you find out
about regressions **before** an extraction job silently degrades. The
recommended audit prompt for your agent harness is at
[`docs/inspection-prompt.md`](docs/inspection-prompt.md).

## What this server does NOT do

- No Microsoft Graph / OAuth — strictly local desktop Outlook
- No `recall_message` (destructive on recipients), no voting buttons (rare)
- No two-phase prepare/confirm pattern — that belongs in the agent prompt
- No "New Outlook" support on either platform until Microsoft restores APIs

## Related projects

- [`Aanerud/outlook-desktop-mcp`](https://github.com/Aanerud/outlook-desktop-mcp)
- [`hasan-imam/mcp-outlook-applescript`](https://github.com/hasan-imam/mcp-outlook-applescript)
- [Model Context Protocol](https://modelcontextprotocol.io)

## License

MIT — see [LICENSE](LICENSE).
