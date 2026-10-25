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
function Invoke-NpmSafe {
    param([string[]]$NpmArgs)
    # npm.cmd writes progress to stderr; under $ErrorActionPreference='Stop'
    # a redirected native stderr would throw spuriously. Run with Continue.
    $prev = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        & $npm @NpmArgs 2>&1 | ForEach-Object { "$_" } | Write-Host
        return $LASTEXITCODE
    } finally {
        $ErrorActionPreference = $prev
    }
}

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
$nodeDir = Split-Path $npm                            # npm.cmd sits in the prefix root, next to node.exe
$node = Join-Path $nodeDir 'node.exe'
if (-not (Test-Path $node)) {
    $node = (Get-Command node -ErrorAction SilentlyContinue).Source
}
if (-not $node -or -not (Test-Path $node)) {
    Write-Host "[installer] Found npm at $npm but no node.exe next to it or on PATH." -ForegroundColor Red
    exit 1
}
Write-Step "Node/npm found: $npm"

# ---------------------------------------------------------------------------
# 2. Install the package globally
# ---------------------------------------------------------------------------

$pkg = 'office-live-mcp'
$installUrl = $null
$prevEap = $ErrorActionPreference
$ErrorActionPreference = 'Continue'
try {
    $prefix = (& $npm prefix -g 2>$null | Where-Object { $_ } | Select-Object -First 1).Trim()
} finally {
    $ErrorActionPreference = $prevEap
}
$installedDir = Join-Path $prefix "node_modules\$pkg"
if (Test-Path (Join-Path $installedDir 'package.json')) {
    Write-Step "$pkg is already installed - updating..."
} else {
    Write-Step "$pkg is not installed yet."
}
Write-Step "Installing $pkg (npm registry first, GitHub fallback)..."
$null = Invoke-NpmSafe @('install', '-g', $pkg)
if (-not (Test-Path (Join-Path $installedDir 'package.json'))) {
    Write-Step "npm registry install failed or unavailable - installing from GitHub..."
    $null = Invoke-NpmSafe @('install', '-g', 'https://github.com/benjaminfrombe/office-live-mcp/archive/refs/heads/main.tar.gz')
    if (-not (Test-Path (Join-Path $installedDir 'package.json'))) {
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

  1. Add the MCP server to your harness (OpenCode, Claude Code or Claude
     Desktop) - the exact steps and config snippets were printed by setup
     just above.
  2. In PowerPoint / Word / Excel: Home -> Add-ins -> Advanced... ->
     SHARED FOLDER tab -> select "Office Live MCP" -> Add.

"@ -ForegroundColor Green
