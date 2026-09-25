$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$templatePath = Join-Path $root 'AppsScript.html'
$outputPath = Join-Path $root 'index.html'
$template = Get-Content -Raw -LiteralPath $templatePath

$includePattern = '<\?!=\s*include\s*\([''\"]([^''\"]+)[''\"]\)\s*\?>'
$staticHtml = [regex]::Replace($template, $includePattern, {
    param($match)
    $relativePath = $match.Groups[1].Value -replace '/', [IO.Path]::DirectorySeparatorChar
    if ([IO.Path]::GetExtension($relativePath) -eq '') {
        $relativePath += '.html'
    }
    $includePath = Join-Path $root $relativePath
    if (-not (Test-Path -LiteralPath $includePath)) {
        throw "Included file not found: $relativePath"
    }
    Get-Content -Raw -LiteralPath $includePath
})

[IO.File]::WriteAllText($outputPath, $staticHtml, [Text.UTF8Encoding]::new($false))
Write-Output "Built $outputPath"