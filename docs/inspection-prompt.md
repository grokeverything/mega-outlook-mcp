# Outlook compatibility inspection

This is a prompt to run periodically (weekly is enough) against your agent
harness so you find out when an Outlook update breaks something **before**
your real extraction job silently degrades.

The MCP exposes an `outlook_diagnostics` tool that does the heavy lifting.
This document is the prompt + scheduling guidance.

## What the prompt does

1. Calls `outlook_diagnostics` (optionally with a known-good
   `sample_message_id` for repeatability).
2. Interprets the structured response.
3. Writes a dated report to disk via `outlook_write_file` if anything is
   not `HEALTHY`.
4. Stops; the harness owns alerting/escalation.

## The prompt

Paste this into your agent harness as a recurring job:

```text
You are auditing the mega-outlook-mcp server against the user's
currently-installed Outlook. Run exactly these steps.

1. Call outlook_diagnostics. The response has:
   - platform               : "windows" | "macos"
   - outlook_version        : the running Outlook version string
   - baseline_version       : the version this MCP was last verified against
   - status                 : "HEALTHY" | "DEGRADED:<n>_fields" | "BROKEN"
   - probed_fields          : { name -> "ok" | "missing" | "error:<msg>" }
   - affected_tools         : list of MCP tool names that depend on a
                              field that is no longer "ok"
   - notes                  : free-form remarks (e.g. "no sample message
                              available")

2. Classify each non-"ok" field:
   - "missing"  : the field exists in our baseline but the API no longer
                  returns it. Treat as a regression — the dependent tools
                  will return null (Windows) or
                  ERROR-MAC-Support-Unavailable (macOS).
   - "error:..." with "permission" / "sandbox" / "not allowed":
                  user is likely on "New Outlook" (sandboxed); recommend
                  switching to classic Outlook.
   - "error:..." with "not running" / "not opened":
                  Outlook is closed; the audit cannot complete. Stop.
   - any other "error:..." :
                  log verbatim, flag for human review.

3. Compare outlook_version against baseline_version:
   - same major.minor          : business as usual
   - one minor behind/ahead    : note "minor drift, monitor"
   - more than one minor ahead : recommend a manual review of the Outlook
                                 AppleScript dictionary (Mac) or the
                                 Outlook Object Model docs (Windows) for
                                 any newly-renamed properties

4. Produce a report (markdown) with these sections:
   - **Summary** : status, outlook_version vs baseline_version
   - **Affected tools** : bullet list (or "none")
   - **Probe details** : the full probed_fields map as a table
   - **Recommended actions** : concrete next steps, ordered by urgency

5. If status is not HEALTHY, write the report to
   `~/outlook-mcp-diagnostics/YYYY-MM-DD.md` using outlook_write_file
   (overwrite=true). Otherwise, do not write anything.

6. Stop. Do not call any other Outlook tools. Do not attempt repairs.
```

## How to run it on a schedule

The MCP itself does not run on a schedule — that's intentional. Use
whatever cron-equivalent your environment already uses.

| Platform | Mechanism                | Recommended cadence |
| -------- | ------------------------ | ------------------- |
| macOS    | `launchd` user agent     | Weekly, Mon 09:00   |
| Windows  | Task Scheduler           | Weekly, Mon 09:00   |
| Cross    | n8n / Temporal / cron    | Weekly              |
| Cloud    | GitHub Actions on a self-hosted runner that has Outlook installed | Weekly |

Ideal timing: after your local Outlook auto-update window. For Microsoft 365
monthly channel, that's typically Patch Tuesday + 2 days.

You can also run `outlook_diagnostics` opportunistically:

- **At server startup**, log the result at INFO. Do not fail-fast on
  DEGRADED — you want the server running so the agent can write its
  diagnostics report.
- **Before a long extraction job**, call once and abort if `BROKEN`.

## What the report tells you to do

| Status                                | Action                                                                                       |
| ------------------------------------- | -------------------------------------------------------------------------------------------- |
| `HEALTHY`                             | Nothing. Discard the report.                                                                 |
| `DEGRADED:<n>_fields`                 | Review affected tools; if you don't use them, ignore. Otherwise file an issue / PR with the renamed field. |
| `BROKEN`                              | Stop trusting extraction output until fixed. Likely causes: New Outlook migration, COM disabled, Outlook not signed in. |

## Updating the baseline

When you intentionally upgrade Outlook (or want to take the new version as
your new known-good):

1. On the upgraded machine, run `outlook_diagnostics` and confirm
   `HEALTHY`.
2. Update `src/mega_outlook_mcp/baseline/outlook_baseline.json`
   with the new `baseline_outlook_versions` string.
3. If any field renamed, update the appropriate template / MAPI tag list
   and re-run diagnostics.
4. Commit.
