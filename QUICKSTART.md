# M-Files Demo Vault Builder — Team Quickstart

Stand up a fully populated M-Files **cloud** demo vault — metadata structure, a
workflow, common views, and (optionally) realistic documents — for **any industry**.
You describe the industry to **Claude Code**; it writes the vault schema YAML; the
`Update-MFilesVault.ps1` tool builds it. Vault-admin rights are enough (no server
system-administrator).

---

## 1. Prerequisites (each machine, one time)

- **Claude Code** (this is how you generate the schema).
- **M-Files Desktop / client** installed (registers the COM API the tool uses).
- **PowerShell 5.1+**.
- **powershell-yaml** module:
  ```powershell
  Install-Module powershell-yaml -Scope CurrentUser
  ```
- The toolkit folder (e.g. `C:\MFilesYamlUpdater`) containing:
  - `Update-MFilesVault.ps1`   ← the build tool
  - `DemoFileFactory.ps1`      ← builds real pdf/docx/xlsx (no Office needed)
  - `New-DemoData.ps1`, `New-HRDemoData.ps1`  ← demo-data generators (real files)
  - `.claude/skills/mfiles-demo-vault/`  ← the Claude Code skill (schema author)
  - `README.md`, `QUICKSTART.md`
- A **cloud vault registered** as a connection in M-Files Desktop, with a login that
  has **"Full control of vault"** (vault admin). Find its name:
  ```powershell
  .\Update-MFilesVault.ps1 -ListConnections
  ```

---

## 2. Build a vault (three moves)

**a) Generate the schema — in Claude Code, from this folder:**

> Open Claude Code in `C:\MFilesYamlUpdater` and say:
> *"Build me a **law firm** demo vault"* (or hospital, manufacturer, insurance, HR…).

The `mfiles-demo-vault` skill writes `<industry>-vault.yaml` for you — value lists,
two object types (parent → child), properties, classes, a workflow, and common
views — all in the tool's exact format.

**b) Apply the structure:**
```powershell
.\Update-MFilesVault.ps1 -YamlPath .\lawfirm-vault.yaml -ApplySchema -SchemaOnly -Cloud
```
No connection/user in the command — it **lists your cloud vaults and prompts for the
vault, username, and password**. (Pass `-ConnectionName "MyVault" -User you@corp.com`
to skip the prompts.)

**c) Add demo data (optional):** ask Claude for a demo-data file, then:
```powershell
.\Update-MFilesVault.ps1 -YamlPath .\lawfirm-demo-data.yaml -Cloud -CreateMissingLookups
```
`-CreateMissingLookups` adds any value-list values the data references. For **bulk
documents with real PDF/Word/Excel files**, use the generator pattern
(`New-DemoData.ps1` / `New-HRDemoData.ps1`).

---

## 3. Good to know

- **Idempotent** — safe to re-run; existing objects are matched and skipped.
- **Vault-admin, not sysadmin** — the big unlock for cloud vaults.
- **Any industry** — the skill designs the whole schema on demand; there are no
  fixed templates to maintain.

---

## 4. Sharing with the team

Zip the folder (or use a Git repo — recommended, so the skill and scripts version
together). **Include** the scripts, the `.claude/` skill folder, and these docs.
**Exclude** generated output:

- `*-vault.yaml`, `*-demo-data.yaml`   (generated per demo)
- `*-demo-files\`                      (generated files)
- `test-*.ps1`                         (diagnostic probes)
