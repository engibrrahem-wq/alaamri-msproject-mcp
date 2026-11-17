@echo off
REM Alaamri MS Project MCP - one-click installer (Developed by Dr Ibrahim M.A. Alaamri, MIT)
REM - Next to server.py  : installs this folder.
REM - On its own         : downloads the latest version from GitHub.
REM Installs into %LOCALAPPDATA%\Alaamri-MSProject-MCP and registers it in Claude Desktop (a .bak backup is kept).
set "AMCP_HERE=%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -Command "iex ((Get-Content -LiteralPath '%~f0' -Raw) -split '#PS-START#')[2]"
echo.
pause
exit /b
#PS-START#
$ErrorActionPreference = 'Stop'
$repoZip = 'https://github.com/engibrrahem-wq/alaamri-msproject-mcp/archive/refs/heads/main.zip'
$here    = $env:AMCP_HERE
$home2   = Join-Path $env:LOCALAPPDATA 'Alaamri-MSProject-MCP'
function Say($m, $c = 'Cyan') { Write-Host $m -ForegroundColor $c }
try {
  Say 'Alaamri MS Project MCP - installer' 'White'
  # 1. Python 3.10+
  $py = $null
  foreach ($cand in @(@('py', '-3'), @('python'))) {
    try { $v = & $cand[0] $cand[1..9] -c "import sys;print('%d.%d'%sys.version_info[:2])" 2>$null; if ($v -and [version]$v -ge [version]'3.10') { $py = $cand; break } } catch {}
  }
  if (-not $py) { throw 'Python 3.10+ not found. Install it from https://www.python.org/downloads/ and tick "Add python.exe to PATH", then run this again.' }
  Say "[1/4] Python $v found"
  # 2. virtual environment
  New-Item -ItemType Directory -Force $home2 | Out-Null
  $venvPy = Join-Path $home2 '.venv\Scripts\python.exe'
  if (-not (Test-Path $venvPy)) { & $py[0] $py[1..9] -m venv (Join-Path $home2 '.venv'); if ($LASTEXITCODE) { throw 'Could not create the virtual environment.' } }
  Say "[2/4] Environment ready: $home2"
  # 3. install (local folder or GitHub)
  $src = if (Test-Path (Join-Path $here 'server.py')) { $here.TrimEnd('\') } else { $repoZip }
  Say "[3/4] Installing from $src (this needs internet) ..."
  & $venvPy -m pip install --disable-pip-version-check -q --upgrade pip | Out-Null
  & $venvPy -m pip install --disable-pip-version-check -q --upgrade --force-reinstall $src
  if ($LASTEXITCODE) { throw 'pip install failed (check your internet connection).' }
  $exe = Join-Path $home2 '.venv\Scripts\alaamri-msproject-mcp.exe'
  if (-not (Test-Path $exe)) { throw "Install finished but $exe is missing." }
  # 4. register in Claude Desktop (normal install + Microsoft Store install), keeping a backup
  $cfgs = @(Join-Path $env:APPDATA 'Claude\claude_desktop_config.json')
  Get-ChildItem (Join-Path $env:LOCALAPPDATA 'Packages') -Directory -Filter 'Claude_*' -ErrorAction SilentlyContinue |
    ForEach-Object { $cfgs += Join-Path $_.FullName 'LocalCache\Roaming\Claude\claude_desktop_config.json' }
  foreach ($c in $cfgs) {
    New-Item -ItemType Directory -Force (Split-Path $c) | Out-Null
    $j = [pscustomobject]@{}
    if (Test-Path $c) {
      Copy-Item $c "$c.bak" -Force
      $raw = Get-Content $c -Raw
      if ($raw -and $raw.Trim()) { $j = $raw | ConvertFrom-Json }
    }
    if (-not ($j.PSObject.Properties.Name -contains 'mcpServers')) { $j | Add-Member -NotePropertyName mcpServers -NotePropertyValue ([pscustomobject]@{}) }
    $j.mcpServers | Add-Member -NotePropertyName 'alaamri-msproject' -NotePropertyValue ([pscustomobject]@{ command = $exe }) -Force
    [IO.File]::WriteAllText($c, ($j | ConvertTo-Json -Depth 32), (New-Object Text.UTF8Encoding $false))
    Say "      registered in $c"
  }
  Say '[4/4] Done.' 'Green'
  Say 'Now FULLY quit Claude Desktop (tray icon > Quit) and open it again.' 'Yellow'
  Say 'Open a COPY of your schedule in MS Project desktop, then ask Claude.' 'Yellow'
} catch {
  Say ('ERROR: ' + $_.Exception.Message) 'Red'
}
