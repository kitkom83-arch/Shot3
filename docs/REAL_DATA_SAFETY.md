# Real Data Safety

This project keeps real dashboard data in `data/`. Smoke runs must never write smoke accounts, jobs, or systems into that directory.

## Safe Smoke Command

Run smoke only through the guarded script:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/smoke-multilane.ps1
```

The smoke script creates a workspace-level `tmp_smoke_runtime_*` folder and sets `SMOKE_DATA_DIR` for the smoke node snippets. Do not run smoke with `SMOKE_DATA_DIR` missing, relative, or pointing at `data/`.

## Required Preflight

Before smoke, run:

```powershell
node --check server.js
powershell -ExecutionPolicy Bypass -File scripts/check-real-data-safety.ps1
```

The safety scan fails if `data/systems` contains smoke markers:

- `acc_smoke_A`
- `acc_smoke_B`
- `smoke_A`
- `smoke_B`
- `tmp_smoke_runtime`
- `selectedAccountId` pointing at a smoke account
- `systemId`, `id`, or `accountId` values starting with `smoke`

When a runtime path is provided, it must be absolute, directly under `tmp_smoke_runtime_*`, end with `data`, and must not equal `workspace/data`.

## Smoke Verification

After smoke, inspect the JSON output:

- `errors` must be `[]`
- `realSmokeHits` must be `[]`
- `safetyScans.before.ok` must be `true`
- `safetyScans.after.ok` must be `true`
- `normal.isolation.selectedAccountsSeparate` must be `true`
- `normal.isolation.aJobDidNotChangeB` must be `true`
- `normal.isolation.bJobDidNotChangeA` must be `true`
- `normal.isolation.autoStatusSeparate` must be `true`
- `normal.isolation.highRetryPausedOnlyA` must be `true`
- `eperm.isolation.epermWarningOnlyA` must be `true`

Do not restart a real server from the smoke script after the run. If a real server is needed, start it manually after confirming the safety scan is clean.

## Manual Safety Scan With Runtime Path

For a specific smoke runtime:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/check-real-data-safety.ps1 -RuntimeDataDir "C:\Users\ADMIN\ยิงเบอร์3-clean\tmp_smoke_runtime_YYYYMMDD_HHMMSS\data"
```

This confirms both the real data tree and the runtime path contract.
