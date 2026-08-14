# M-Files YAML Vault Builder

Build an **M-Files (on-prem)** vault — metadata structure *and* objects — from a
single YAML file, via the M-Files COM API in PowerShell.

`Update-MFilesVault.ps1` does two things:

1. **Schema** (`-ApplySchema`) — creates value lists (+items), object types,
   property definitions, and classes. Runs over an administrative
   (Windows-integrated) connection. Alias-driven and **idempotent** — re-running
   skips what exists and only fills gaps.
2. **Objects** — creates the records in the `objects:` section. Runs over an
   **early-bound** connection as a named M-Files user (`-User`, or you're
   prompted for it — so different runs can use different accounts).

## Quick start with Claude Code (recommended)

This repo ships a Claude Code **skill** (`.claude/skills/mfiles-demo-vault/`) that
writes a vault schema YAML for any industry on demand — you don't hand-author YAML.

1. Clone this repo, then **open the repo folder as your workspace/project root in
   Claude Code** — in the desktop app, start the chat *in* the
   `mfiles-demo-vault-builder` folder (open/select it as the workspace). Do **not**
   open a different folder and add this one as a secondary directory — Claude Code
   only discovers skills from the **project root's** `.claude\` folder, so a folder
   added on the side won't expose the skill.
2. Type `/mfiles-demo-vault` (it should appear in the `/` list), or just ask:
   *"build me a law firm demo vault"* (hospital, manufacturer, insurance, HR, …).
3. The skill writes `<industry>-vault.yaml` into the folder and hands you the exact
   `Update-MFilesVault.ps1` build command to run against your cloud vault.

> If `/mfiles-demo-vault` doesn't show up, your session root isn't this folder —
> re-open the repo folder as the workspace root and start a fresh chat (skills are
> indexed at session start). Alternatively, install it once for every session by
> copying `.claude\skills\mfiles-demo-vault` into your user-level `~\.claude\skills\`.

Everything below is the underlying tool the skill drives — read on if you want to
author YAML by hand or understand what the build command does.

## Requirements

- Windows with the **M-Files client/API** installed (registers the COM library).
- **PowerShell 5.1+**.
- **`powershell-yaml`** module: `Install-Module powershell-yaml -Scope CurrentUser`
- **`Interop.MFilesAPI.dll`** in this folder (needed for object creation — the
  early-bound path). Only required when creating objects; `-SchemaOnly` doesn't need it.
- An M-Files login (pass `-User`, or enter it at the prompt) that is a **vault
  user with create rights** (and a **vault admin** if you're applying schema).

## Usage

```powershell
# List registered connections (to find names/GUIDs)
.\Update-MFilesVault.ps1 -ListConnections

# Structure only (safe to re-run; no object duplicates)
.\Update-MFilesVault.ps1 -YamlPath .\accounting-firm.yaml -ApplySchema -SchemaOnly

# Full build: structure, then objects (prompts for the user, then password)
.\Update-MFilesVault.ps1 -YamlPath .\accounting-firm.yaml -ApplySchema

# Objects only (structure already exists)
.\Update-MFilesVault.ps1 -YamlPath .\accounting-firm.yaml

# Validate without writing
.\Update-MFilesVault.ps1 -YamlPath .\accounting-firm.yaml -ApplySchema -DryRun
```

Useful switches: `-User <login>` (object identity), `-Interop <path>`,
`-SchemaPhase valuelists|objecttypes|properties|classes`, `-CreateMissingLookups`,
`-AllowDuplicates` (always create objects, skip the existence check),
`-SingleLogin` (use one M-Files-user connection for both schema and objects).

```powershell
# Single login: use the -User account (must be a vault admin) for everything
.\Update-MFilesVault.ps1 -YamlPath .\accounting-firm.yaml -ApplySchema -SingleLogin

# Cloud vault: connect via the registered connection's own endpoint (grpc/TLS),
# as a vault-admin M-Files user - no server system-administrator needed
.\Update-MFilesVault.ps1 -YamlPath .\credit-union.yaml -ApplySchema -Cloud `
    -ConnectionName "Credit Union" -User andyn@demoemails.com
```

### Cloud vaults

`-Cloud` (alias of `-UseConnectionEndpoint`) makes the tool connect through the
**registered connection's endpoint** (protocol, address, port, TLS) instead of
`localhost`, so it reaches an M-Files **Cloud** vault. It logs in administratively
as `-User`, so that account needs **"Full control of vault"** (vault admin) — you
do **not** need server System Administrator. Auth defaults to M-Files-user
(`-AuthType 3`); use `-AuthType 1` for Windows/SSO-backed logins.

## YAML format

See [`accounting-firm.yaml`](accounting-firm.yaml). Sections are applied in order.

```yaml
connection:
  name: ClaudeTest          # registered connection name (resolves the vault GUID)
  # vault: "{GUID}"         # or specify the GUID directly

value_lists:
  - name: Service Line
    items: [Tax, Audit, Advisory]

object_types:
  - name: Client
    name_plural: Clients

property_definitions:
  - { name: EIN,     type: text }
  - { name: Amount,  type: number }
  - { name: Client,  type: multilookup, based_on: Client }   # lookup to an object type
  - { name: Preparer,type: lookup,      based_on: Users }    # lookup to built-in Users

classes:
  - name: Tax Return
    object_type: Document           # built-in Document type, or a custom object type name
    properties: [Client, Amount, EIN]

workflows:
  - name: Engagement Workflow
    states: [Draft, In Review, Approved, Filed]
    transitions:
      - { from: No State,  to: Draft }      # 'No State' = the entry point
      - { from: Draft,     to: In Review }
      - { from: In Review, to: Approved }
      - { from: Approved,  to: Filed }

views:
  - name: Documents by Client
    group_by: [Client]                      # one folder level per property
  - name: Documents by Client and Fiscal Year
    group_by: [Client, Fiscal Year]         # nested folders
    # common: true                          # default; false = personal view

objects:
  - class: Client
    title: "Acme Manufacturing LLC"
    properties:
      EIN: "12-3456789"
    files:
      - "C:\\path\\to\\file.pdf"     # optional attachments (documents)
```

Property `type` values: `text, multiline, integer, number, date, timestamp,
boolean, lookup, multilookup`. Lookups need `based_on` (a value-list name, an
object-type name, or `Users`). Object-type lookups adopt the auto-created lookup
property instead of duplicating it.

**Workflows** (`workflows:`) are created in the schema phase. Each has `states:`
(a list of names) and `transitions:` (`from`/`to` referencing state names).
`No State` is M-Files' implicit start — a transition `from: No State` is the
entry point. Workflows, states, and transitions are alias-driven and idempotent:
re-running adds only what's missing (aliases `WF.*`, `WFS.*`, `WFT.*`).

**Views** (`views:`) create shared **common** views (schema phase). `group_by`
is a list of property names — each becomes a folder level, so `[Client, Fiscal
Year]` nests Fiscal-Year folders under each Client. `common: false` makes it a
personal view. Views have no alias, so they're matched by **name** (re-running
skips existing ones).

## How it works / notes

- **Two connections on purpose:** structure changes need a server-admin
  connection; object creation needs a vault user *and* early binding (M-Files'
  `CreateNewObjectEx` has an optional argument that COM *late* binding can't
  satisfy — hence the interop assembly).
- **"Users can create objects of this type"** is enabled automatically on object
  types the tool creates (and repaired on existing ones), so object creation
  isn't blocked by permissions.
- **Objects are idempotent.** Each object is matched against the vault — by an
  explicit `match:` block, otherwise by its **title** within the class's object
  type. Found → **updated** (check-out / set properties / check-in); not found →
  **created**. So the whole tool is safely re-runnable and converges to the YAML.
  Use `id:` for a direct update by internal ID, or `-AllowDuplicates` to force
  plain creates.
- **No-op runs don't bump the version.** Before updating, the object's current
  property values (and files, if `files:` is set) are compared to the YAML; if
  they already match, the object is reported `Unchanged` and left as-is — no
  check-out, no new version. (Files are compared by count and size.)
- Start against a **test vault** and use `-DryRun` first.

- **Files on update are synced to match `files:`** — on an update, the object's
  files are replaced with exactly the list in `files:` (existing files removed,
  the specified ones added). Omit `files:` to leave an object's files untouched.

- **Single-login mode** (`-SingleLogin`): the two connections exist because
  structure needs server-admin rights and object creation needs a vault user +
  early binding. When your `-User` account is a vault **admin**, `-SingleLogin`
  collapses both into one connection (one credential prompt).
