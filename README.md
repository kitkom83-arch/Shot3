# Telegram Checker Dashboard

## Safe Smoke Run

Use the guarded multi-lane smoke script only:

```powershell
node --check server.js
powershell -ExecutionPolicy Bypass -File scripts/check-real-data-safety.ps1
powershell -ExecutionPolicy Bypass -File scripts/smoke-multilane.ps1
```

Smoke data must stay under a workspace-level `tmp_smoke_runtime_*` directory. Do not run smoke with `SMOKE_DATA_DIR` missing, relative, or pointing at `data/`.

After smoke, verify the JSON output has `realSmokeHits: []`, `errors: []`, `safetyScans.before.ok: true`, `safetyScans.after.ok: true`, and the A/B isolation checks set to `true`.

Restore and backup recovery steps are documented in `docs/BACKUP_RECOVERY.md`. Real data safety rules are documented in `docs/REAL_DATA_SAFETY.md`.
