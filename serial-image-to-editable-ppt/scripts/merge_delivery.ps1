[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$OutputDir,

    [int]$Start = 1,

    [Parameter(Mandatory = $true)]
    [int]$End,

    [string]$OutputFile,

    [string]$PreviewDir,

    [switch]$Force
)

$ErrorActionPreference = 'Stop'

if ($Start -gt $End) {
    throw 'Merge requires ascending order: Start must be <= End.'
}

$resolvedOutputDir = (Resolve-Path -LiteralPath $OutputDir).Path
if (-not $OutputFile) {
    $OutputFile = Join-Path $resolvedOutputDir "pages_${Start}_to_${End}_merged_editable.pptx"
}
$OutputFile = [System.IO.Path]::GetFullPath($OutputFile)
$outputParent = Split-Path -Parent $OutputFile
if (-not (Test-Path -LiteralPath $outputParent)) {
    New-Item -ItemType Directory -Path $outputParent -Force | Out-Null
}

if ((Test-Path -LiteralPath $OutputFile) -and -not $Force) {
    throw "Output already exists. Use -Force to replace it: $OutputFile"
}

$sources = foreach ($page in $Start..$End) {
    $source = Join-Path $resolvedOutputDir "page_${page}_refined_editable.pptx"
    if (-not (Test-Path -LiteralPath $source)) {
        throw "Missing page deck: $source"
    }
    $source
}

if ($Force -and (Test-Path -LiteralPath $OutputFile)) {
    Remove-Item -LiteralPath $OutputFile -Force
}

$previewPath = $null
if ($PreviewDir) {
    $previewPath = [System.IO.Path]::GetFullPath($PreviewDir)
    New-Item -ItemType Directory -Path $previewPath -Force | Out-Null
    if ((Get-ChildItem -LiteralPath $previewPath -Filter '*.PNG' -ErrorAction SilentlyContinue) -and -not $Force) {
        throw "Preview directory contains PNG files. Use -Force to replace them: $previewPath"
    }
    if ($Force) {
        Get-ChildItem -LiteralPath $previewPath -Filter '*.PNG' -ErrorAction SilentlyContinue |
            Remove-Item -Force
    }
}

$application = New-Object -ComObject PowerPoint.Application
$presentation = $null
try {
    $presentation = $application.Presentations.Add()
    foreach ($source in $sources) {
        $inserted = $presentation.Slides.InsertFromFile($source, $presentation.Slides.Count)
        if ($inserted -ne 1) {
            throw "Expected one slide from $source; inserted $inserted"
        }
    }

    $expected = $End - $Start + 1
    if ($presentation.Slides.Count -ne $expected) {
        throw "Merged slide count mismatch: expected $expected, got $($presentation.Slides.Count)"
    }

    $presentation.SaveAs($OutputFile, 24)
    if ($previewPath) {
        $presentation.Export($previewPath, 'PNG', 1672, 941)
    }
    $slideCount = $presentation.Slides.Count
    $presentation.Close()
    $presentation = $null
}
finally {
    if ($presentation) {
        $presentation.Close()
    }
    $application.Quit()
    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($application)
}

$previewCount = if ($previewPath) {
    (Get-ChildItem -LiteralPath $previewPath -Filter '*.PNG').Count
}
else {
    0
}

[pscustomobject]@{
    output = $OutputFile
    start = $Start
    end = $End
    slides = $slideCount
    preview_dir = $previewPath
    previews = $previewCount
    bytes = (Get-Item -LiteralPath $OutputFile).Length
} | ConvertTo-Json -Compress
