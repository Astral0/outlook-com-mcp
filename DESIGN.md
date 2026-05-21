# Design notes — why COM Automation?

This document explains the constraints that led to building yet another Outlook MCP server, and why it talks to Outlook through COM Automation rather than through Microsoft Graph, EWS, or IMAP.

## TL;DR

In an unmodified enterprise tenant, **Microsoft Graph delegated** is the right way to drive a mailbox from an LLM. This project exists because some tenants close every supported path:

- App registrations are refused by IT;
- First-party clients (Azure CLI, Microsoft Graph PowerShell, Teams) are removed, assignment-required, or have their preauth path blocked by Microsoft;
- EWS still answers TCP but only via OAuth (same admin gate as Graph);
- IMAP/SMTP are blocked at the firewall *and* deprecated by Microsoft since 2022.

In that environment, the only surface left is the **Outlook desktop client itself**, already SSO-authenticated by the user. COM Automation is the most stable, scriptable, and unambiguous way to talk to it. Hence this server.

If your tenant lets you register an app and consent to `Mail.Read` / `Calendars.ReadWrite`, **stop reading and go use Graph.**

## The four options evaluated

Measured in April 2026 on a real corporate workstation in a heavily-locked enterprise tenant.

| Option | Verdict | What happened |
|---|---|---|
| **Microsoft Graph (delegated)** | ❌ Blocked at tenant | App registration request refused. First-party fallbacks all closed (see table below). |
| **EWS via `exchangelib`** | ⚠️ Functionally blocked | EWS endpoint reachable (HTTP 401). Microsoft retired EWS Basic Auth in October 2022 — OAuth-only — which routes back through the same admin gate as Graph. |
| **IMAP/SMTP** | ❌ Dead end | Ports 993/587 timeout at the corporate firewall. Microsoft retired IMAP/SMTP Basic Auth in September 2022 anyway. |
| **Playwright on OWA** | ✅ Viable fallback | `outlook.office.com/mail/` loads with transparent SSO via the user's Edge session. Fragile and slow for batch work, but unblockable. |
| **COM via `pywin32`** | ✅ **Chosen** | Outlook M365 16.0.x reachable through `GetActiveObject("Outlook.Application")`. Direct store access. No tenant permission. No popup on save/delete operations in normal flows. Send was validated under explicit user confirmation. |

## Why Graph is blocked in this tenant

Tested via MSAL against three Microsoft-first-party public clients:

| Client | AADSTS code | Mechanism |
|---|---|---|
| Azure CLI | `700016` | Application **removed from the tenant** by admin. |
| Microsoft Teams | `65002` | **Microsoft itself** blocks the Teams→Graph preauth path (anti-spoofing measure against first-party impersonation). Teams can only call the Teams Service API, not generic Graph. |
| Microsoft Graph PowerShell | `50105` | App is present but **assignment-required**; only users explicitly granted by admin can use it. The test user is not granted. |

Three independent doors, three different failure modes. There is no end-user workaround in this specific tenant. Your tenant may differ — re-run the equivalent test before assuming Graph is unreachable.

## Why COM was chosen over Playwright

Playwright on OWA also works and was kept as a fallback option, but COM beats it on every operational axis:

- **Speed**: COM is in-process. Playwright re-renders the OWA UI for each action.
- **Stability**: COM uses MAPI primitives that have been stable for two decades. OWA's DOM changes without notice.
- **Determinism**: COM exposes `EntryID`s (stable identifiers). Scraping OWA depends on element selectors that drift.
- **Attachments**: COM gives direct access to attachment binaries. Playwright requires triggering a download and reading from disk.
- **Batch operations**: COM handles thousands of items per call. Playwright would time out or get rate-limited by OWA's lazy-loading.

The trade-off: COM requires Outlook Classic to be running on a Windows machine. That is fine in a corporate context where Outlook is always open; it would be unworkable in a serverless context.

## Architecture choices

- **Single-file Python server** using FastMCP over stdio. ~2000 lines, no plugin system. Easy to audit, easy to fork.
- **Two-step write model**: every state-changing operation either creates a draft (visible in Outlook's Drafts folder) or requires an explicit `confirm=True` flag. Re-validation of the recipient allowlist happens at send time, not only at draft creation, so a draft modified by hand cannot bypass the allowlist.
- **Recipient allowlist**: `OUTLOOK_MCP_ALLOWED_DOMAINS` defaults to empty (no restriction). Setting it to your corporate domain prevents the LLM from accidentally emailing the outside world during exploration — this is strongly recommended in any non-trivial deployment.
- **X.500 → SMTP resolution**: in-tenant senders appear as `/O=EXCHANGELABS/OU=...` rather than as SMTP addresses. The server resolves SMTP via `PropertyAccessor` (`PR_SMTP_ADDRESS = 0x39FE001E`) when possible, so the allowlist actually works on the addresses you expect.
- **Coverage transparency**: list/search responses include `truncated: bool` and `coverage: {oldest_received, newest_received}` so the caller (typically an LLM) can tell whether its temporal window was fully covered or silently capped. This was added after the LLM produced confident analyses of "the last three weeks" while only seeing the last eight days.
- **`since_iso` uses DASL syntax**: an early version mixed Jet brackets (`[ReceivedTime]`) with `@SQL=...` URIs, which Outlook silently rejects. The correct form is `"urn:schemas:httpmail:datereceived" >= 'YYYY-MM-DDTHH:MM:SS'`.

## What is deliberately *not* in scope

- **Shared mailboxes** other than the default profile. Possible but unimplemented.
- **GAL Free/Busy** for other people's calendars. Possible via `Recipient.FreeBusy` but flaky; not exposed.
- **Rules engine**: a partial rules surface exists (`create_rule`, `list_rules`, `toggle_rule`, `delete_rule`) but Outlook's RuleConditions API is famously incomplete and silently downgrades unsupported conditions. Use at your own risk.
- **Cross-platform support**. COM is Windows-only. macOS Outlook AppleScript could be a different project.
- **"New Outlook"** (the WebView2-based UI). It exposes no COM. Users must run Outlook Classic.

## When this server would become obsolete

Any of the following changes would make this approach unnecessary, in roughly increasing order of friction:

1. Tenant admin assigns the user to Microsoft Graph PowerShell → standard MSAL device-code flow works → use a Graph-based MCP instead.
2. Tenant accepts a third-party multi-tenant Graph app (e.g. a vendor-supplied MCP) → same outcome.
3. Tenant grants a custom app registration with `Mail.ReadWrite` / `Calendars.ReadWrite` delegated permissions → ideal case.

Re-run a minimal MSAL device-code probe against your tenant before pivoting. Graph remains the architecturally correct answer wherever it is reachable.

## Related work

This server fills a gap that the existing Outlook MCP ecosystem does not address: every public Outlook MCP server I found at the time of writing requires Graph credentials or assumes the user has admin consent for an app. None of them attempt COM Automation against the local desktop client.

If a better Graph-based server is available **and your tenant allows it**, use that. This project is not trying to compete with Graph — it is trying to keep working when Graph is closed.
