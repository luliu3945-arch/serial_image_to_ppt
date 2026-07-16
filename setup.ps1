[CmdletBinding()]
param(
    [string]$SkillRoot = (Join-Path $HOME '.agents\skills'),
    [string]$RuntimeRoot = (Join-Path $HOME '.serial-image-to-ppt'),
    [string]$PythonCommand = 'python',
    [switch]$SkipDependencies
)

$ErrorActionPreference = 'Stop'
$RepoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$SkillNames = @('serial-image-to-editable-ppt', 'codeximage-to-editable-ppt-v1')

Write-Host '1/5 Checking repository files...'
foreach ($name in $SkillNames) {
    $skillFile = Join-Path $RepoRoot "$name\SKILL.md"
    if (-not (Test-Path -LiteralPath $skillFile)) {
        throw "Missing skill file: $skillFile"
    }
}

Write-Host '2/5 Checking Python 3.10+...'
$versionText = & $PythonCommand -c "import sys; print('.'.join(map(str, sys.version_info[:3]))); raise SystemExit(0 if sys.version_info >= (3,10) else 1)"
if ($LASTEXITCODE -ne 0) {
    throw "Python 3.10 or newer is required. Detected: $versionText"
}

$VenvDir = Join-Path $RuntimeRoot 'venv'
$VenvPython = Join-Path $VenvDir 'Scripts\python.exe'
if (-not $SkipDependencies) {
    Write-Host "3/5 Creating/updating isolated runtime at $VenvDir..."
    New-Item -ItemType Directory -Path $RuntimeRoot -Force | Out-Null
    if (-not (Test-Path -LiteralPath $VenvPython)) {
        & $PythonCommand -m venv $VenvDir
        if ($LASTEXITCODE -ne 0) { throw 'Failed to create the Python virtual environment.' }
    }
    & $VenvPython -m pip install --upgrade pip
    if ($LASTEXITCODE -ne 0) { throw 'Failed to upgrade pip.' }
    & $VenvPython -m pip install -r (Join-Path $RepoRoot 'requirements.txt')
    if ($LASTEXITCODE -ne 0) { throw 'Failed to install Python dependencies.' }
}
elseif (-not (Test-Path -LiteralPath $VenvPython)) {
    $VenvPython = (Get-Command $PythonCommand).Source
}

Write-Host "4/5 Installing both skills to $SkillRoot..."
New-Item -ItemType Directory -Path $SkillRoot -Force | Out-Null
foreach ($name in $SkillNames) {
    $source = Join-Path $RepoRoot $name
    $destination = Join-Path $SkillRoot $name
    New-Item -ItemType Directory -Path $destination -Force | Out-Null
    Copy-Item -Path (Join-Path $source '*') -Destination $destination -Recurse -Force
}

$runtimeConfig = @{
    python = [System.IO.Path]::GetFullPath($VenvPython)
    repository = [System.IO.Path]::GetFullPath($RepoRoot)
    installed_at = (Get-Date).ToString('o')
} | ConvertTo-Json
$runtimeConfigPath = Join-Path $SkillRoot 'serial-image-to-editable-ppt\runtime.json'
[System.IO.File]::WriteAllText($runtimeConfigPath, $runtimeConfig, [System.Text.UTF8Encoding]::new($false))

Write-Host '5/5 Running environment check...'
& $VenvPython (Join-Path $RepoRoot 'scripts\doctor.py') --strict --skill-root $SkillRoot
if ($LASTEXITCODE -ne 0) { throw 'Core environment validation failed.' }

Write-Host ''
Write-Host 'Installation complete.' -ForegroundColor Green
Write-Host "Skills: $SkillRoot"
Write-Host "Python runtime: $VenvPython"
Write-Host 'Restart Codex if the two skills do not appear immediately.'
Write-Host 'Optional production checks:'
Write-Host "  & '$VenvPython' '$RepoRoot\scripts\doctor.py' --require-powerpoint --require-ocr"
