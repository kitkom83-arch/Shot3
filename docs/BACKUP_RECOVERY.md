# Backup Recovery Checklist

Use this checklist before and after smoke runs or manual restore work. Do not delete real backups during smoke cleanup.

## Before Running Smoke

Confirm these are backed up or preserved:

- `data/` as a full snapshot under `data_backups/data_backup_*`
- existing `data_before_restore_*` folders
- the latest known-good backup name and timestamp
- smoke JSON output from the last run, if it is needed for audit

The guarded smoke script creates `data_backups/data_backup_<timestamp>` before it starts the smoke server.

## Restore From `data_backups`

Use a timestamped backup such as `data_backups/data_backup_20260511_174126`.

```powershell
Rename-Item -LiteralPath data -NewName data_before_restore_YYYYMMDD_HHMMSS
Copy-Item -LiteralPath data_backups\data_backup_20260511_174126 -Destination data -Recurse
```

Keep the moved `data_before_restore_*` folder until the restored dashboard has been inspected and a safety scan passes.

## Verify Restore

Run:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/check-real-data-safety.ps1
node --check server.js
```

Then inspect:

- `data/systems/A` through `data/systems/E` exist as expected
- `accounts.json`, `app_state.json`, and `settings.json` are present for restored systems
- no `acc_smoke_A`, `acc_smoke_B`, `smoke_A`, `smoke_B`, or `tmp_smoke_runtime` markers appear under `data/systems`
- selected real accounts in `app_state.json` are not smoke account ids

## Verify Smoke Did Not Mix Into Real Data

After any smoke run:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/check-real-data-safety.ps1
```

The command must return `ok: true` and `smokeHits: []`. The smoke script JSON must also show `realSmokeHits: []`.

## Keep `data_before_restore_*`

Keep the moved folder until all of these are true:

- the restore source path is recorded
- the safety scan passes
- the dashboard opens with the expected systems and accounts
- no operator needs files from the pre-restore copy
- at least one newer backup has been created after the restore

Only delete an old `data_before_restore_*` folder after confirming it is not the only copy of any needed real data.

## When Old Backups Can Be Deleted

Delete old backups only after:

- a newer backup has been verified with the safety scan
- the restored or current `data/` has been used successfully
- there is no unresolved incident requiring comparison against that backup
- the backup is not the latest known-good restore point

Never delete backups as part of smoke cleanup. Smoke cleanup may remove `tmp_smoke_runtime_*` folders only after their output has been reviewed.
