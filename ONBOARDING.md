# M-Files Demo Vault Builder — SE Onboarding

Stand up a fully-populated M-Files **cloud** demo vault for any industry, straight
from Claude Code. You describe the industry; the `mfiles-demo-vault` skill writes the
schema; the `Update-MFilesVault.ps1` tool builds it. **Vault-admin rights are enough —
no server system-administrator needed.**

---

## 1. Prerequisites (one time per machine)

- **Claude Code** (desktop app or CLI)
- **M-Files Desktop client** — registers the COM API the tool uses, and is how you
  register the cloud vaults you want to build
- **PowerShell 5.1+** (built into Windows)
- **powershell-yaml** module:
  ```powershell
  Install-Module powershell-yaml -Scope CurrentUser
  ```
- **Allow local scripts** (once):
  ```powershell
  Set-ExecutionPolicy -Scope CurrentUser RemoteSigned -Force
  ```

## 2. Clone the repo

```bash
git clone https://github.com/jdedwards10-inspire/mfiles-demo-vault-builder.git
```

(It's a **private** repo — if the clone is denied, ask to be added as a collaborator.)

## 3. Open it in Claude Code — the one detail that matters

**Open the cloned `mfiles-demo-vault-builder` folder as your workspace/project ROOT**
in a **new** Claude Code chat — *not* a parent folder with the repo nested inside it.
Skills load only from the root folder's `.claude/`, and only at session start.

**Verify:** type `/` — you should see **`mfiles-demo-vault`** in the list. If you
don't, your session root isn't the repo folder — reopen the repo folder as the root
and start a fresh chat.

## 4. Register your demo vault

In M-Files Desktop, log into the cloud vault you want to build (once) so the tool can
see it. Your account needs **"Full control of vault"** (vault admin) on it — *not*
server System Administrator.

## 5. Build a vault

In the repo folder, in Claude Code:

1. Say **"build me a \<industry\> demo vault"** (law firm, healthcare, manufacturing,
   insurance, HR, real estate, …), or type `/mfiles-demo-vault`. It writes
   `<industry>-vault.yaml`.
2. Run the build command it gives you — no editing needed:
   ```powershell
   .\Update-MFilesVault.ps1 -YamlPath .\<industry>-vault.yaml -ApplySchema -SchemaOnly -Cloud
   ```
   It lists your registered vaults and **prompts for the vault, username, and
   password**.
3. *(Optional)* Ask Claude for a demo-data file, then load it with
   `-CreateMissingLookups`:
   ```powershell
   .\Update-MFilesVault.ps1 -YamlPath .\<industry>-demo-data.yaml -Cloud -CreateMissingLookups
   ```

## Good to know

- **Idempotent** — safe to re-run; existing objects are matched and skipped.
- **Any industry** — the skill designs the whole schema on demand; there are no fixed
  templates to maintain.
- **Cloud-first** — works against M-Files Cloud with just vault-admin rights.
- **`git pull`** keeps the skill and scripts current.
