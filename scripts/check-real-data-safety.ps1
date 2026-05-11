param(
  [string]$RuntimeDataDir = $env:SMOKE_DATA_DIR,
  [string]$Workspace = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
)

$ErrorActionPreference = "Stop"

function Resolve-FullPath {
  param([string]$Path)
  return [System.IO.Path]::GetFullPath($Path)
}

function Add-Hit {
  param(
    [System.Collections.Generic.List[object]]$Hits,
    [string]$Type,
    [string]$Path,
    [string]$Detail
  )
  $Hits.Add([ordered]@{
    type = $Type
    path = $Path
    detail = $Detail
  }) | Out-Null
}

$workspacePath = Resolve-FullPath $Workspace
$realDataDir = Resolve-FullPath (Join-Path $workspacePath "data")
$systemsDir = Join-Path $realDataDir "systems"
$hits = [System.Collections.Generic.List[object]]::new()

if ($RuntimeDataDir) {
  if (-not [System.IO.Path]::IsPathRooted($RuntimeDataDir)) {
    Add-Hit $hits "runtimeDataDir" $RuntimeDataDir "runtimeDataDir must be absolute"
  } else {
    $runtimeDataPath = Resolve-FullPath $RuntimeDataDir
    $runtimeRoot = Split-Path -Parent $runtimeDataPath
    $runtimeRootName = Split-Path -Leaf $runtimeRoot

    if ($runtimeDataPath -eq $realDataDir) {
      Add-Hit $hits "runtimeDataDir" $runtimeDataPath "runtimeDataDir must not equal workspace/data"
    }
    if ($runtimeRootName -notlike "tmp_smoke_runtime_*") {
      Add-Hit $hits "runtimeDataDir" $runtimeDataPath "runtimeDataDir must be directly under tmp_smoke_runtime_*"
    }
    if (-not ($runtimeDataPath.StartsWith($runtimeRoot + [System.IO.Path]::DirectorySeparatorChar))) {
      Add-Hit $hits "runtimeDataDir" $runtimeDataPath "runtimeDataDir must be inside tmp_smoke_runtime_*"
    }
    if ($runtimeDataPath -notmatch "tmp_smoke_runtime_") {
      Add-Hit $hits "runtimeDataDir" $runtimeDataPath "runtimeDataDir must contain tmp_smoke_runtime_"
    }
    if ((Split-Path -Leaf $runtimeDataPath) -ne "data") {
      Add-Hit $hits "runtimeDataDir" $runtimeDataPath "runtimeDataDir must end with data"
    }
  }
}

if (Test-Path $realDataDir) {
  Get-ChildItem -LiteralPath $realDataDir -Directory -Filter "tmp_smoke_runtime*" -ErrorAction SilentlyContinue |
    ForEach-Object {
      Add-Hit $hits "realDataPath" $_.FullName "tmp_smoke_runtime directory must not exist under workspace/data"
    }
}

if (Test-Path $systemsDir) {
  Get-ChildItem -LiteralPath $systemsDir -Directory -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -match "^(?i)smoke" } |
    ForEach-Object {
      Add-Hit $hits "systemId" $_.FullName "system id must not start with smoke"
    }

  $patterns = @(
    'acc_smoke_A',
    'acc_smoke_B',
    'smoke_A',
    'smoke_B',
    'tmp_smoke_runtime',
    '"selectedAccountId"\s*:\s*"(acc_smoke_|smoke)',
    '"(id|accountId)"\s*:\s*"smoke',
    '"(id|accountId)"\s*:\s*"acc_smoke_',
    '"systemId"\s*:\s*"smoke',
    '\bSmoke A\b',
    '\bSmoke B\b'
  )
  $pattern = $patterns -join "|"
  $scanHits = rg -n --hidden --glob "!*.xlsx" --glob "!*.png" --glob "!*.jpg" --glob "!*.jpeg" --glob "!*.gif" --glob "!*.zip" $pattern $systemsDir 2>$null
  if ($LASTEXITCODE -eq 0) {
    foreach ($line in $scanHits) {
      Add-Hit $hits "smokeMarker" $systemsDir $line
    }
  } elseif ($LASTEXITCODE -gt 1) {
    Add-Hit $hits "scanError" $systemsDir "rg failed with exit code $LASTEXITCODE"
  }
}

$result = [ordered]@{
  ok = ($hits.Count -eq 0)
  workspace = $workspacePath
  realDataDir = $realDataDir
  systemsDir = $systemsDir
  runtimeDataDir = $RuntimeDataDir
  smokeHits = @($hits)
}

$result | ConvertTo-Json -Depth 20
if ($hits.Count -gt 0) {
  exit 1
}
