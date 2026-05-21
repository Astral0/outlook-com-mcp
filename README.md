# outlook-com-mcp

> **MCP server for Microsoft Outlook desktop via COM (pywin32) — for locked-down corporate environments where Microsoft Graph is not an option.**

This is a [Model Context Protocol](https://modelcontextprotocol.io) server that exposes a Windows Outlook (M365 desktop) mailbox to LLM clients such as Claude Code, Cursor, or Gemini CLI. It drives the **already-authenticated Outlook desktop client** through COM Automation, so it inherits the user's SSO session and requires no tenant permissions, no app registration, and no admin involvement.

It exists specifically for the case where Microsoft Graph delegated access is blocked by tenant policy (assignment-required apps, first-party preauth blocks, Azure CLI uninstalled, etc.). If your tenant allows Graph, **use Graph instead** — it is architecturally and operationally superior. See [DESIGN.md](./DESIGN.md) for the full rationale.

---

## ⚠️ Status: personal project, "as is"

This server is built and used by one person in one specific corporate environment. It works on my machine. It is published as-is, with no support, no roadmap, no guarantee, and no commitment of any kind. **Use at your own risk.**

Read the code before running it. Understand what it touches. Test in a controlled mailbox first. The server can read, write, send, and modify your Outlook data within the limits documented below.

---

## What it does

Exposes 27 MCP tools across five areas:

- **Mail (read)**: `whoami`, `list_folders`, `list_mail`, `read_mail`, `search_mail`, `download_attachment`, `health_check`
- **Mail (write, guarded)**: `create_draft`, `send_draft`, `reply_mail`, `move_mail`, `mark_read`, `flag_mail`, `guardrails_status`
- **Calendar**: `list_events`, `read_event`, `find_freeslots`, `find_freeslots_multi`, `create_event_draft`, `send_event_invites`, `respond_meeting`
- **Rules**: `list_rules`, `toggle_rule`, `create_rule`, `delete_rule`
- **Contacts (GAL via LDAP)**: `search_contacts`, `get_contact_details`

All operations run against the Outlook profile of the user who launched Outlook. There is no separate authentication step.

## Requirements

- **Windows** (10/11). The COM bridge is Windows-only.
- **Outlook Classic** desktop (M365 or Office 2019+). The "New Outlook" UI and Outlook on the web are **not supported** — they expose no COM surface.
- **Python 3.11+** with `pywin32` and the `mcp` SDK.
- A live Outlook session — the server connects to the running instance via `GetActiveObject("Outlook.Application")` (it will fall back to `Dispatch` and launch Outlook if not running).

## Install

Not published on PyPI. Install from source:

```bash
git clone https://github.com/Astral0/outlook-com-mcp.git
cd outlook-com-mcp
pip install -e .
```

You can also just clone the repo and run `python -m outlook_com_mcp` from inside it without installing, as long as `pywin32` and `mcp` are on your `PYTHONPATH`.

## Install with an AI agent (recommended)

This project assumes you already use an LLM coding agent (Claude Code, Cursor, etc.). The fastest, most reliable way to install and adapt it to *your* corporate context is to let the agent do the work. After cloning the repo, open it in your agent and paste a prompt like this:

> Read `README.md` and `DESIGN.md` in this repo. Then help me install and configure this MCP server for my environment. Specifically:
>
> 1. Detect my Python environment (interpreter path, version, whether `pywin32` and `mcp` are already installed). Install missing dependencies.
> 2. Verify that Outlook Classic desktop is running on this machine. Try a minimal COM smoke test (e.g. opening `Outlook.Application` via `pywin32` and reading `Application.DefaultProfileName`). Stop if Outlook is not reachable and explain why.
> 3. Ask me for my corporate email domain(s) so we can set `OUTLOOK_MCP_ALLOWED_DOMAINS` correctly. Default to *blocking sends to the outside world* — never weaken this without explicit confirmation.
> 4. Locate my MCP client's config file (e.g. `~/.claude/mcp.json`, `~/.cursor/mcp.json`, `~/.gemini/settings.json`) and add an `outlook` server entry using the absolute path to the right Python interpreter. Include `PYTHONIOENCODING=utf-8`.
> 5. After config, instruct me to restart the MCP client and then call the `whoami` tool to confirm the server is alive. Then call `list_mail(limit=3)` to confirm read access works.
> 6. Read the **Known limitations and gotchas** section of `README.md` and warn me about anything that applies to my setup.
>
> Do not send any email or modify any calendar event during install. Drafts and read-only calls only.

The agent will detect your context, fail loudly on missing prerequisites, configure the right MCP entry, and validate end-to-end. This is also a useful pattern for keeping the install reproducible: store the prompt you actually used in your team's onboarding docs.

## Configure your MCP client

Minimal `.mcp.json` entry (Claude Code, Cursor, etc.):

```json
{
  "mcpServers": {
    "outlook": {
      "command": "python",
      "args": ["-m", "outlook_com_mcp"],
      "env": {
        "PYTHONIOENCODING": "utf-8",
        "OUTLOOK_MCP_ALLOWED_DOMAINS": "",
        "OUTLOOK_MCP_ALLOW_SEND": "0"
      }
    }
  }
}
```

`PYTHONIOENCODING=utf-8` is mandatory on Windows — otherwise the default `cp1252` mangles accented characters in tool output.

See [`examples/mcp.json`](./examples/mcp.json) for variants (Gemini CLI, restricted domains, etc.).

## Safety: write guardrails

The server defaults to a conservative, two-step write model:

| Variable | Default | Meaning |
|---|---|---|
| `OUTLOOK_MCP_ALLOW_SEND` | `0` | If `0`, `reply_mail(save_only=False)` and `send_event_invites(send=True)` are refused. Drafts are still created. Set to `1` only if you trust the LLM to send mail unattended. |
| `OUTLOOK_MCP_ALLOWED_DOMAINS` | *(empty)* | Comma-separated allowlist of recipient domains. Empty means no restriction. Recommended to restrict in corporate contexts (e.g. `"yourcorp.com"`). |

Recommended workflow for write actions:

1. Tool creates a draft (`create_draft` / `reply_mail` with `save_only=True`).
2. You review it in Outlook's Drafts folder.
3. You click **Send** manually, **or** call `send_draft(entry_id, confirm=True)`.

Calling `send_draft` revalidates the recipient allowlist at send time, even if the draft was modified after creation.

## Logging

`%LOCALAPPDATA%\outlook-com-mcp\server.log` (rotating, 5 MB × 5).

## Known limitations and gotchas

- `Items.Restrict("@SQL=...")` requires DASL URIs throughout. Mixing Jet-style `[ReceivedTime]` with `@SQL=` is rejected by Outlook (silent gotcha).
- `IncludeRecurrences = True` + `Items.Restrict()` returns zero results — a documented Outlook bug. `find_freeslots` iterates directly with an early-break instead.
- `find_freeslots` only inspects **your own calendar**. GAL Free/Busy lookup is not implemented (would require a different API surface).
- `sender_address` returned by Outlook is in X.500 format (`/O=EXCHANGELABS/...`) for in-tenant senders. The server resolves SMTP via `PropertyAccessor` (`PR_SMTP_ADDRESS`) when possible.
- `EntryID` is the only stable identifier. Subject, index, and conversation ID are **not** stable across moves/sessions.
- `search_contacts` uses ADODB/LDAP against the AD that the workstation is joined to. If the host is not domain-joined, it falls back to Outlook's `CreateRecipient` ANR (single result).
- No support for shared mailboxes other than the default profile.
- Outlook prompts (Object Model Guard, modal dialogs) will block the COM bridge silently. Run with Outlook in a normal interactive session.

## License

MIT — see [`LICENSE`](./LICENSE).

## Acknowledgements

Inspired by the constraints of working in a heavily-locked tenant. See [DESIGN.md](./DESIGN.md) for the full story.
