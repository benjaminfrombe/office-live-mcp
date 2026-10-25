# Office Live MCP — Windows installer
# One-liner:
#   irm https://raw.githubusercontent.com/benjaminfrombe/office-live-mcp/main/install.ps1 | iex
#
# What it does:
#   1. Finds Node.js + npm (PATH, C:\node\*, Program Files, nvm-windows, volta)
#      — or offers to install Node LTS via winget when nothing is found
#   2. Installs the office-live-mcp package globally (npm registry, GitHub fallback)
#   3. Runs `office-live-mcp setup` (TLS cert, trusted catalog, registry)
#   4. Prints the one remaining manual step (Office requires it)

$ErrorActionPreference = 'Stop'

function Write-Step([string]$msg) { Write-Host "[installer] $msg" -ForegroundColor Cyan }

# ---------------------------------------------------------------------------
# 1. Locate Node + npm
# ---------------------------------------------------------------------------

function Find-Npm {
    foreach ($candidate in @(
        (Get-Command npm.cmd -ErrorAction SilentlyContinue),
        (Get-Command npm -ErrorAction SilentlyContinue)
    )) {
        if ($candidate) { return $candidate.Source }
    }
    $locations = @(
        "$env:SystemDrive\node\*\npm.cmd",
        "$env:ProgramFiles\nodejs\npm.cmd",
        "${env:ProgramFiles(x86)}\nodejs\npm.cmd",
        "$env:LOCALAPPDATA\Programs\nodejs\npm.cmd",   # nodejs.org user installer
        "$env:NVM_HOME\npm.cmd",                        # nvm-windows
        "$env:LOCALAPPDATA\Volta\bin\npm.cmd",          # volta
        "$env:USERPROFILE\scoop\apps\nodejs*\current\npm.cmd",
        "$env:ProgramData\chocolatey\bin\npm.exe"       # chocolatey shim
    )
    foreach ($pattern in $locations) {
        $found = Get-Item $pattern -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($found) { return $found.FullName }
    }
    return $null
}

$npm = Find-Npm
if (-not $npm) {
    Write-Host "[installer] Node.js not found." -ForegroundColor Yellow
    $winget = Get-Command winget -ErrorAction SilentlyContinue
    if ($winget) {
        $answer = Read-Host "Install Node.js LTS with winget? (y/n)"
        if ($answer -eq 'y' -or $answer -eq 'Y') {
            Write-Step "Installing Node.js LTS via winget..."
            winget install --id OpenJS.NodeJS.LTS --accept-source-agreements --accept-package-agreements
            # Refresh PATH for this session, then retry.
            $env:Path = [Environment]::GetEnvironmentVariable('Path', 'Machine') + ';' + [Environment]::GetEnvironmentVariable('Path', 'User')
            $npm = Find-Npm
        }
    }
}
if (-not $npm) {
    Write-Host @"

[installer] Cannot continue without Node.js. Install it from https://nodejs.org
           (LTS is fine), open a NEW terminal, and run this installer again.
"@ -ForegroundColor Red
    exit 1
}
$nodeDir = Split-Path (Split-Path $npm)   # npm.cmd lives in <prefix>, node.exe next to it
$node = Join-Path $nodeDir 'node.exe'
if (-not (Test-Path $node)) { $node = (Get-Command node -ErrorAction SilentlyContinue).Source }
Write-Step "Node/npm found: $npm"

# ---------------------------------------------------------------------------
# 2. Install the package globally
# ---------------------------------------------------------------------------

$pkg = 'office-live-mcp'
$installUrl = $null
$prefix = & $npm prefix -g
$installedDir = Join-Path $prefix "node_modules\$pkg"
if (Test-Path (Join-Path $installedDir 'package.json')) {
    Write-Step "$pkg is already installed - updating..."
} else {
    Write-Step "$pkg is not installed yet."
}
Write-Step "Installing $pkg (npm registry first, GitHub fallback)..."
& $npm install -g $pkg 2>$null
if ($LASTEXITCODE -ne 0 -or -not (Test-Path (Join-Path $installedDir 'package.json'))) {
    Write-Step "npm registry install failed or unavailable - installing from GitHub..."
    & $npm install -g 'https://github.com/benjaminfrombe/office-live-mcp/archive/refs/heads/main.tar.gz'
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path (Join-Path $installedDir 'package.json'))) {
        Write-Host "[installer] Installation failed. See the npm output above." -ForegroundColor Red
        exit 1
    }
}
Write-Step "Installed to $installedDir"

# ---------------------------------------------------------------------------
# 3. Run setup
# ---------------------------------------------------------------------------

Write-Step "Running setup..."
$entry = Join-Path $installedDir 'dist\index.cjs'
& $node $entry setup
if ($LASTEXITCODE -ne 0) {
    Write-Host "[installer] setup reported a failure - see the messages above." -ForegroundColor Red
    exit 1
}

# ---------------------------------------------------------------------------
# 4. Next steps
# ---------------------------------------------------------------------------

Write-Host @"

[installer] Done. Two things left:

  1. The bridge daemon starts automatically the first time an MCP client
     (Claude Code, OpenCode, Claude Desktop) connects.
  2. In PowerPoint / Word / Excel: Home -> Add-ins -> Advanced... ->
     MY ORGANIZATION tab -> select "Office Live MCP" -> Add.

"@ -ForegroundColor Green
