param(
  [int]$Port = 3001
)

$ErrorActionPreference = "Stop"

$Workspace = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$ServerPath = Join-Path $Workspace "server.js"
$SafetyScript = Join-Path $PSScriptRoot "check-real-data-safety.ps1"
$RealDataDir = Join-Path $Workspace "data"
$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$RuntimeDir = Join-Path $Workspace "tmp_smoke_runtime_$Timestamp"
$RuntimeDataDir = Join-Path $RuntimeDir "data"
$BackupRoot = Join-Path $Workspace "data_backups"
$BackupDir = Join-Path $BackupRoot "data_backup_$Timestamp"

function Resolve-FullPath {
  param([string]$Path)
  return [System.IO.Path]::GetFullPath($Path)
}

function Assert-RuntimeDataDir {
  $workspaceData = Resolve-FullPath $RealDataDir
  $runtimeRoot = Resolve-FullPath $RuntimeDir
  $runtimeData = Resolve-FullPath $RuntimeDataDir

  if (-not [System.IO.Path]::IsPathRooted($runtimeData)) {
    throw "runtime DATA_DIR must be absolute: $runtimeData"
  }
  if ($runtimeData -eq $workspaceData) {
    throw "runtime DATA_DIR must not be workspace/data: $runtimeData"
  }
  if (-not (Split-Path -Leaf $runtimeRoot).StartsWith("tmp_smoke_runtime_")) {
    throw "runtime root must be tmp_smoke_runtime_*: $runtimeRoot"
  }
  if (-not ($runtimeData.StartsWith($runtimeRoot + [System.IO.Path]::DirectorySeparatorChar))) {
    throw "runtime DATA_DIR must be inside tmp_smoke_runtime_*: $runtimeData"
  }
  if ($runtimeData -notmatch "tmp_smoke_runtime_") {
    throw "runtime DATA_DIR must contain tmp_smoke_runtime: $runtimeData"
  }
}

function Assert-Condition {
  param(
    [bool]$Condition,
    [string]$Message
  )
  if (-not $Condition) { throw $Message }
}

function Get-ListenPid {
  param([int]$LocalPort)
  $pattern = "0\.0\.0\.0:$LocalPort\s+0\.0\.0\.0:0\s+LISTENING"
  $line = netstat -ano | Select-String $pattern | Select-Object -First 1
  if (-not $line) { return $null }
  $parts = $line.ToString().Trim() -split "\s+"
  return [int]$parts[-1]
}

function Wait-Health {
  param([int]$LocalPort)
  $deadline = (Get-Date).AddSeconds(25)
  do {
    try {
      $response = Invoke-WebRequest -UseBasicParsing "http://127.0.0.1:$LocalPort/api/health" -TimeoutSec 2
      if ($response.StatusCode -eq 200) { return }
    } catch {
      Start-Sleep -Milliseconds 500
    }
  } while ((Get-Date) -lt $deadline)
  throw "server on port $LocalPort did not become healthy"
}

function Write-SmokeEnv {
  param([bool]$SimulateEperm)

  $envText = Get-Content (Join-Path $Workspace ".env") -Raw
  $envText = $envText -replace "(?m)^PORT=.*$", "PORT=$Port"
  $envText = $envText -replace "(?m)^DATA_DIR=.*$", "DATA_DIR=$RuntimeDataDir"
  if ($envText -notmatch "(?m)^DATA_DIR=") {
    $envText += "`nDATA_DIR=$RuntimeDataDir`n"
  }
  if ($SimulateEperm) {
    $envText += "`nSIMULATE_STATE_RENAME_EPERM=current_job.json`nSIMULATE_STATE_RENAME_EPERM_COUNT=4`n"
  }
  Set-Content -Path (Join-Path $RuntimeDir ".env") -Value $envText -Encoding UTF8
}

function Start-SmokeServer {
  param([bool]$SimulateEperm)

  Write-SmokeEnv -SimulateEperm $SimulateEperm
  $stdout = Join-Path $RuntimeDir ($(if ($SimulateEperm) { "server-eperm.stdout.log" } else { "server.stdout.log" }))
  $stderr = Join-Path $RuntimeDir ($(if ($SimulateEperm) { "server-eperm.stderr.log" } else { "server.stderr.log" }))
  $process = Start-Process -FilePath node -ArgumentList $ServerPath -WorkingDirectory $RuntimeDir -RedirectStandardOutput $stdout -RedirectStandardError $stderr -PassThru -WindowStyle Hidden
  Wait-Health -LocalPort $Port
  return $process
}

function Assert-NoSmokeInRealData {
  $scan = & powershell -ExecutionPolicy Bypass -File $SafetyScript -RuntimeDataDir $RuntimeDataDir -Workspace $Workspace | ConvertFrom-Json
  if ($LASTEXITCODE -ne 0 -or $scan.ok -ne $true) {
    return @($scan.smokeHits)
  }
  return @()
}

function Invoke-RealDataSafety {
  param([string]$Phase)

  $scan = & powershell -ExecutionPolicy Bypass -File $SafetyScript -RuntimeDataDir $RuntimeDataDir -Workspace $Workspace | ConvertFrom-Json
  if ($LASTEXITCODE -ne 0 -or $scan.ok -ne $true) {
    throw "real data safety scan failed during ${Phase}: $($scan.smokeHits | ConvertTo-Json -Compress -Depth 12)"
  }
  return $scan
}

function Backup-RealData {
  New-Item -ItemType Directory -Force -Path $BackupRoot | Out-Null
  if (Test-Path $RealDataDir) {
    Copy-Item -LiteralPath $RealDataDir -Destination $BackupDir -Recurse -Force
  } else {
    New-Item -ItemType Directory -Force -Path $BackupDir | Out-Null
  }
  return $BackupDir
}

function Invoke-NodeSmoke {
  param([string]$Script)

  $oldSmokeDataDir = $env:SMOKE_DATA_DIR
  try {
    $env:SMOKE_DATA_DIR = $RuntimeDataDir
    Push-Location $RuntimeDir
    try {
      return ($Script | node -)
    } finally {
      Pop-Location
    }
  } finally {
    $env:SMOKE_DATA_DIR = $oldSmokeDataDir
  }
}

$result = [ordered]@{
  workspace = $Workspace
  runtime = $RuntimeDir
  runtimeDataDir = $RuntimeDataDir
  backup = $null
  stoppedPid = $null
  nodeCheck = $false
  normal = $null
  eperm = $null
  realSmokeHits = @()
  safetyScans = [ordered]@{
    before = $null
    after = $null
  }
  errors = @()
}

$server = $null

try {
  Assert-RuntimeDataDir
  $result.safetyScans.before = Invoke-RealDataSafety -Phase "before-smoke"

  $result.backup = Backup-RealData
  New-Item -ItemType Directory -Force -Path $RuntimeDataDir | Out-Null
  Assert-RuntimeDataDir

  node --check $ServerPath | Out-Null
  $result.nodeCheck = $true

  $existingPid = Get-ListenPid -LocalPort $Port
  if ($existingPid) {
    $result.stoppedPid = $existingPid
    Stop-Process -Id $existingPid -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 800
  }

  $server = Start-SmokeServer -SimulateEperm $false

  $normalScript = @'
(async () => {
  const fs = require("fs");
  const fsp = require("fs/promises");
  const path = require("path");
  const crypto = require("crypto");

  const base = "http://127.0.0.1:3001";
  function getSmokeDataDir() {
    const value = process.env.SMOKE_DATA_DIR;
    if (!value) throw new Error("SMOKE_DATA_DIR is required");
    if (!path.isAbsolute(value)) throw new Error("SMOKE_DATA_DIR must be absolute");
    const resolved = path.resolve(value);
    if (!resolved.includes("tmp_smoke_runtime_")) throw new Error("SMOKE_DATA_DIR must point at tmp_smoke_runtime data");
    if (path.basename(path.dirname(resolved)).startsWith("tmp_smoke_runtime_") === false) {
      throw new Error("SMOKE_DATA_DIR must be directly under tmp_smoke_runtime_*");
    }
    if (path.basename(resolved) !== "data") throw new Error("SMOKE_DATA_DIR must end with data");
    return resolved;
  }
  const dataDir = getSmokeDataDir();

  let cookie = "";
  const result = { api: {}, isolation: {}, tempDataSystems: [], details: { dataDir } };

  async function req(method, url, body) {
    const headers = { cookie };
    let payload;
    if (body !== undefined) {
      headers["content-type"] = "application/json";
      payload = JSON.stringify(body);
    }
    const res = await fetch(base + url, { method, headers, body: payload });
    const text = await res.text();
    let data = {};
    try { data = text ? JSON.parse(text) : {}; } catch { data = { raw: text }; }
    return { status: res.status, ok: res.ok, data };
  }

  async function login() {
    const res = await fetch(base + "/auth/login", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ username: "admin", password: "Admin_2026_StrongPass_9f4K" }),
    });
    cookie = (res.headers.get("set-cookie") || "").split(";")[0];
    if (!res.ok || !cookie) throw new Error(`login failed ${res.status}`);
  }

  async function writeJson(file, data) {
    const resolvedFile = path.resolve(file);
    if (!resolvedFile.includes("tmp_smoke_runtime_")) throw new Error(`smoke writes require tmp_smoke_runtime path: ${file}`);
    if (!resolvedFile.startsWith(path.resolve(dataDir) + path.sep)) throw new Error(`write outside smoke data: ${file}`);
    await fsp.mkdir(path.dirname(file), { recursive: true });
    await fsp.writeFile(file, JSON.stringify(data, null, 2));
  }

  async function writeText(file, text) {
    const resolvedFile = path.resolve(file);
    if (!resolvedFile.includes("tmp_smoke_runtime_")) throw new Error(`smoke writes require tmp_smoke_runtime path: ${file}`);
    if (!resolvedFile.startsWith(path.resolve(dataDir) + path.sep)) throw new Error(`write outside smoke data: ${file}`);
    await fsp.mkdir(path.dirname(file), { recursive: true });
    await fsp.writeFile(file, text, "utf8");
  }

  function account(systemId) {
    return [{
      id: `acc_smoke_${systemId}`,
      label: `Smoke ${systemId}`,
      apiId: "12345",
      apiHashEnc: "invalid",
      phone: systemId === "A" ? "+66000000001" : "+66000000002",
      sessionEnc: "smoke-session",
      status: "connected",
      pendingPhoneCodeHash: "",
      pendingCodeType: "",
      pendingTimeout: 0,
      awaitingPassword: false,
      lastError: "",
      lastUsedAt: "",
      me: null,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    }];
  }

  async function prepareSystem(systemId) {
    const root = path.join(dataDir, "systems", systemId);
    await writeJson(path.join(root, "accounts.json"), account(systemId));
    await writeJson(path.join(root, "app_state.json"), { selectedAccountId: `acc_smoke_${systemId}` });
    await writeJson(path.join(root, "settings.json"), {
      autoRun: false,
      batchSize: 1,
      maxContactsPerRun: 20,
      delayBetweenRunsSec: 120,
      retryPauseSec: 600,
      retryRatioThreshold: 0.2,
      waitFloodAutomatically: true,
    });
    await writeText(
      path.join(root, "input", "clean", "clean_ready.csv"),
      `name,phone,rawPhone,sourceLabel,consentStatus,customerStatus\nSmoke ${systemId} One,+660000001${systemId.charCodeAt(0)},+660000001${systemId.charCodeAt(0)},smoke,yes,new\nSmoke ${systemId} Two,+660000002${systemId.charCodeAt(0)},+660000002${systemId.charCodeAt(0)},smoke,yes,buyer\n`
    );
    await writeJson(path.join(root, "input", "clean", "intake_latest.json"), {
      fileName: `smoke-${systemId}.csv`,
      jobLabel: `Smoke ${systemId}`,
      jobNote: "smoke",
      totalRows: 2,
      readyRows: 2,
    });
  }

  async function manifest(dir) {
    const out = [];
    async function walk(current) {
      if (!fs.existsSync(current)) return;
      const entries = await fsp.readdir(current, { withFileTypes: true });
      entries.sort((a, b) => a.name.localeCompare(b.name));
      for (const entry of entries) {
        const fullPath = path.join(current, entry.name);
        if (entry.isDirectory()) await walk(fullPath);
        else {
          const buf = await fsp.readFile(fullPath);
          out.push(`${path.relative(dir, fullPath).replaceAll("\\", "/")}:${crypto.createHash("sha256").update(buf).digest("hex")}`);
        }
      }
    }
    await walk(dir);
    return out.join("\n");
  }

  async function setJobHighRetry(systemId) {
    const out = path.join(dataDir, "systems", systemId, "output");
    for (const name of ["current_job.json", "job_state.json"]) {
      const file = path.join(out, name);
      const job = JSON.parse(await fsp.readFile(file, "utf8"));
      job.status = "ready";
      job.autoStatus = "RUNNING";
      job.lastRetryRatio = 0.85;
      job.manualRetryRequired = false;
      job.nextRunAt = "";
      job.updatedAt = new Date().toISOString();
      await writeJson(file, job);
    }
  }

  await login();
  result.api.health = (await fetch(base + "/api/health")).status;
  const systems = await req("GET", "/api/systems");
  result.api.systems = systems.status;
  result.api.systemIds = (systems.data.systems || []).map((item) => item.id);
  result.api.statusA = (await req("GET", "/api/systems/A/status")).status;
  result.api.statusB = (await req("GET", "/api/systems/B/status")).status;

  await prepareSystem("A");
  await prepareSystem("B");

  const accA = await req("GET", "/api/systems/A/accounts");
  const accB = await req("GET", "/api/systems/B/accounts");
  result.isolation.selectedAccountsSeparate =
    accA.data.accounts?.[0]?.selected === true &&
    accB.data.accounts?.[0]?.selected === true &&
    accA.data.accounts?.[0]?.id !== accB.data.accounts?.[0]?.id;

  const bBeforeAJob = await manifest(path.join(dataDir, "systems", "B"));
  const createA = await req("POST", "/api/systems/A/job/create");
  const bAfterAJob = await manifest(path.join(dataDir, "systems", "B"));
  result.api.createJobA = createA.status;
  result.isolation.aJobDidNotChangeB = bBeforeAJob === bAfterAJob;

  const aBeforeBJob = await manifest(path.join(dataDir, "systems", "A"));
  const createB = await req("POST", "/api/systems/B/job/create");
  const aAfterBJob = await manifest(path.join(dataDir, "systems", "A"));
  result.api.createJobB = createB.status;
  result.isolation.bJobDidNotChangeA = aBeforeBJob === aAfterBJob;

  const aJobPath = path.join(dataDir, "systems", "A", "output", "current_job.json");
  const bJobPath = path.join(dataDir, "systems", "B", "output", "current_job.json");
  const aJobStatePath = path.join(dataDir, "systems", "A", "output", "job_state.json");
  const bJobStatePath = path.join(dataDir, "systems", "B", "output", "job_state.json");

  const aJob = JSON.parse(await fsp.readFile(aJobPath, "utf8"));
  aJob.autoStatus = "RUNNING";
  aJob.status = "ready";
  await writeJson(aJobPath, aJob);
  await writeJson(aJobStatePath, aJob);

  const bJob = JSON.parse(await fsp.readFile(bJobPath, "utf8"));
  bJob.autoStatus = "OFF";
  bJob.status = "ready";
  await writeJson(bJobPath, bJob);
  await writeJson(bJobStatePath, bJob);

  const autoA = await req("GET", "/api/systems/A/auto-status");
  const autoB = await req("GET", "/api/systems/B/auto-status");
  result.isolation.autoStatusSeparate = autoA.data.status === "RUNNING" && autoB.data.status === "OFF";
  result.details.autoA = autoA.data.status;
  result.details.autoB = autoB.data.status;

  const bBeforeRetry = await manifest(path.join(dataDir, "systems", "B"));
  await setJobHighRetry("A");
  const highRetryA = await req("GET", "/api/systems/A/auto-status");
  const bAfterRetry = await manifest(path.join(dataDir, "systems", "B"));
  const highRetryB = await req("GET", "/api/systems/B/auto-status");
  result.api.highRetryA = highRetryA.status;
  result.isolation.highRetryPausedOnlyA =
    highRetryA.data.queueStatus === "PAUSED_TOO_MANY_RETRY" &&
    highRetryA.data.status === "OFF" &&
    highRetryB.data.queueStatus !== "PAUSED_TOO_MANY_RETRY" &&
    bBeforeRetry === bAfterRetry;
  result.details.highRetryAQueue = highRetryA.data.queueStatus;
  result.details.highRetryBQueue = highRetryB.data.queueStatus;

  result.tempDataSystems = (await fsp.readdir(path.join(dataDir, "systems"), { withFileTypes: true }))
    .filter((entry) => entry.isDirectory())
    .map((entry) => entry.name)
    .sort();

  process.stdout.write(JSON.stringify(result));
})().catch((error) => {
  console.error(error.stack || error.message);
  process.exit(1);
});
'@

  $result.normal = (Invoke-NodeSmoke -Script $normalScript) | ConvertFrom-Json
  Assert-Condition ($result.normal.isolation.selectedAccountsSeparate -eq $true) "A/B selected account isolation failed"
  Assert-Condition ($result.normal.isolation.aJobDidNotChangeB -eq $true) "A job changed B data"
  Assert-Condition ($result.normal.isolation.bJobDidNotChangeA -eq $true) "B job changed A data"
  Assert-Condition ($result.normal.isolation.autoStatusSeparate -eq $true) "A/B auto status isolation failed"
  Assert-Condition ($result.normal.isolation.highRetryPausedOnlyA -eq $true) "high retry pause did not stay isolated to A"

  Stop-Process -Id $server.Id -ErrorAction SilentlyContinue
  $server = $null
  Start-Sleep -Milliseconds 800

  $server = Start-SmokeServer -SimulateEperm $true

  $epermScript = @'
(async () => {
  const fsp = require("fs/promises");
  const path = require("path");
  const base = "http://127.0.0.1:3001";
  function getSmokeDataDir() {
    const value = process.env.SMOKE_DATA_DIR;
    if (!value) throw new Error("SMOKE_DATA_DIR is required");
    if (!path.isAbsolute(value)) throw new Error("SMOKE_DATA_DIR must be absolute");
    const resolved = path.resolve(value);
    if (!resolved.includes("tmp_smoke_runtime_")) throw new Error("SMOKE_DATA_DIR must point at tmp_smoke_runtime data");
    if (path.basename(path.dirname(resolved)).startsWith("tmp_smoke_runtime_") === false) {
      throw new Error("SMOKE_DATA_DIR must be directly under tmp_smoke_runtime_*");
    }
    if (path.basename(resolved) !== "data") throw new Error("SMOKE_DATA_DIR must end with data");
    return resolved;
  }
  const dataDir = getSmokeDataDir();

  let cookie = "";
  async function req(method, url, body) {
    const headers = { cookie };
    let payload;
    if (body !== undefined) {
      headers["content-type"] = "application/json";
      payload = JSON.stringify(body);
    }
    const res = await fetch(base + url, { method, headers, body: payload });
    const text = await res.text();
    let data = {};
    try { data = text ? JSON.parse(text) : {}; } catch { data = { raw: text }; }
    return { status: res.status, ok: res.ok, data };
  }
  async function writeJson(file, data) {
    const resolvedFile = path.resolve(file);
    if (!resolvedFile.includes("tmp_smoke_runtime_")) throw new Error(`smoke writes require tmp_smoke_runtime path: ${file}`);
    if (!resolvedFile.startsWith(path.resolve(dataDir) + path.sep)) throw new Error(`write outside smoke data: ${file}`);
    await fsp.mkdir(path.dirname(file), { recursive: true });
    await fsp.writeFile(file, JSON.stringify(data, null, 2));
  }
  async function login() {
    const res = await fetch(base + "/auth/login", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ username: "admin", password: "Admin_2026_StrongPass_9f4K" }),
    });
    cookie = (res.headers.get("set-cookie") || "").split(";")[0];
    if (!res.ok || !cookie) throw new Error(`login failed ${res.status}`);
  }
  async function setHigh(systemId) {
    const out = path.join(dataDir, "systems", systemId, "output");
    for (const name of ["current_job.json", "job_state.json"]) {
      const file = path.join(out, name);
      const job = JSON.parse(await fsp.readFile(file, "utf8"));
      job.status = "ready";
      job.autoStatus = "RUNNING";
      job.lastRetryRatio = 0.85;
      job.manualRetryRequired = false;
      job.nextRunAt = "";
      job.updatedAt = new Date().toISOString();
      await writeJson(file, job);
    }
  }

  await login();
  await setHigh("A");
  const aAuto = await req("GET", "/api/systems/A/auto-status");
  const aDash = await req("GET", "/api/systems/A/dashboard");
  const bDash = await req("GET", "/api/systems/B/dashboard");
  const aWarnings = (aDash.data.recovery || []).map((item) => item.code);
  const bWarnings = (bDash.data.recovery || []).map((item) => item.code);
  process.stdout.write(JSON.stringify({
    api: {
      epermAStatus: aAuto.status,
      epermADashboard: aDash.status,
      epermBDashboard: bDash.status,
    },
    isolation: {
      epermWarningOnlyA: aWarnings.includes("STATE_WRITE_FAILED") && !bWarnings.includes("STATE_WRITE_FAILED"),
    },
    details: {
      aWarnings,
      bWarnings,
      simulationTarget: "current_job.json",
    },
  }));
})().catch((error) => {
  console.error(error.stack || error.message);
  process.exit(1);
});
'@

  $result.eperm = (Invoke-NodeSmoke -Script $epermScript) | ConvertFrom-Json
  Assert-Condition ($result.eperm.isolation.epermWarningOnlyA -eq $true) "EPERM warning was not isolated to A"
} catch {
  $result.errors += $_.Exception.Message
} finally {
  if ($server -and (Get-Process -Id $server.Id -ErrorAction SilentlyContinue)) {
    Stop-Process -Id $server.Id -ErrorAction SilentlyContinue
  }
  Start-Sleep -Milliseconds 800
  try {
    $result.safetyScans.after = Invoke-RealDataSafety -Phase "after-smoke"
    $result.realSmokeHits = @()
  } catch {
    $result.errors += $_.Exception.Message
    $result.realSmokeHits = @(Assert-NoSmokeInRealData)
  }
}

$result | ConvertTo-Json -Depth 30
if ($result.realSmokeHits.Count -gt 0) {
  exit 1
}
if ($result.errors.Count -gt 0) {
  exit 1
}
