# excel365-mcp: a live Excel 365 connector for Claude

Lets Claude read and edit a shared Excel workbook on SharePoint or OneDrive **while your team keeps working in it**. Edits land in the live file within seconds, the same as a co-author typing.

It was built for finance work: trial balances, notes to accounts, financial statements, inventory valuation. That means:
- **Edits are guarded.** An edit is skipped if a teammate changed the cell since Claude read it.
- **Every change is logged and can be undone.**
- **Tools for auditing a workbook:**
  - where errors start;
  - circular references;
  - which trial-balance rows feed a SUMIFS;
  - formulas with typed-in amounts.

It's a Claude Code plugin with an MCP server (Model Context Protocol, the standard way to give Claude new tools). The server talks to the Microsoft Graph Excel API, using only the Python standard library plus `openpyxl`.

## Install (Claude Code on Windows)
1. **Add the plugin:**
   ```
   /plugin marketplace add uday-sreenivasula/excel365-mcp
   /plugin install excel365@excel365-mcp
   ```
2. **Start a new Claude Code session.** The first start takes 1–3 minutes: it checks for Python 3.10+ (installing Python 3.12 for your user with `winget` if it's missing), then installs `openpyxl` into `%LOCALAPPDATA%\excel365\lib`. If Claude says the connector failed to start, wait a minute and start a new session.
3. **Ask Claude to open a workbook** by pasting its SharePoint/OneDrive link or giving the file name. **Microsoft's sign-in page opens in your browser.** Sign in with your work account and you're done.

**Your password stays with Microsoft.** You sign in on Microsoft's own page, with MFA, through Microsoft's own "Microsoft Graph Command Line Tools" app. Claude can only open files your account can already open, and edits show under your name.

**Admin approval:** some organisations require an admin to approve "Microsoft Graph Command Line Tools" (delegated `Files.ReadWrite.All`) once for the tenant. If sign-in says "Need admin approval", ask your Microsoft 365 admin.

**Restricting sign-in to one organisation (optional):** set the environment variable `EXCEL365_TENANT` to your tenant domain, e.g. `contoso.com`. By default, any Microsoft 365 work or school account can sign in.

## What Claude can do with it
| Tool | Purpose |
|---|---|
| `xl_open` | Open a workbook by link or file name, and list its sheets (with `details`, also used ranges and formula counts) |
| `xl_map` | A sheet's structure: copied formulas collapsed to one line, hardcoded numbers, labels, formulas with typed-in amounts |
| `xl_read` | Read live cells (formula and value), several areas in one call |
| `xl_find` | Search formulas/values, with total counts and per-sheet grouping |
| `xl_trace` | What a cell is built from, including which rows a SUMIFS/XLOOKUP actually picks up, and what uses it |
| `xl_write` / `xl_fill` | Edit cells with a conflict guard. Batches are all-or-nothing; `as_text` keeps date-like text as text; `watch` returns recalculated totals |
| `xl_undo` / `xl_log` | Undo this chat's edits; an audit trail with old values; per-call timings |
| `xl_changes` | What anyone changed since a checkpoint |
| `xl_check` | Where errors start (and how many cells each breaks), circular references, `#REF!` formulas, external links |
| `xl_insert` / `xl_delete` / `xl_format` | Insert or delete cells, rows and columns; formatting across many areas at once |
| `xl_sheet` / `xl_add_sheet` / `xl_delete_sheet` | Hide, unhide, rename, reorder, add and delete sheets |
| `xl_api` | Direct access to the Microsoft Graph API for anything else: named ranges, tables, charts, protection, recalculation |

Each Claude chat opens its own workbook, and that chat remembers it, even after Claude restarts the connector. Tools never act on a workbook another chat opened. To work across two workbooks in one chat, add `workbook=<file name>` to any tool.

### Skip the permission prompts (optional)
To let Claude use these tools without asking each time, add this to `%USERPROFILE%\.claude\settings.json` under `permissions.allow`:
```json
"mcp__plugin_excel365_excel365"
```

## Requirements and troubleshooting
- **Windows only.** Sign-in is stored encrypted for your Windows user (DPAPI) in `%USERPROFILE%\.xlgraph`.
- **Desktop Excel users should keep AutoSave on.** Ask people who have the file open in desktop Excel to save with AutoSave on (or close it) before Claude inserts or deletes rows. Otherwise their copy may fail to merge.
- **Map, find, trace and check use a downloaded snapshot**, which can be a minute or two behind live edits. Reads and writes are always live.
- **If the connection to Microsoft stops responding**, the connector stops after one timed-out call and says so, instead of hanging.
- **Setup log:** `%LOCALAPPDATA%\excel365\setup.log`. To force setup to run again, delete `%LOCALAPPDATA%\excel365\ready-1`.
- **To sign out**, delete `%USERPROFILE%\.xlgraph\refresh.dat`.

## Licence
MIT. See [LICENSE](LICENSE).
