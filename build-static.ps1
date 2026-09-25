$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$templatePath = Join-Path $root 'AppsScript.html'
$outputPath = Join-Path $root 'index.html'
$utf8 = [Text.UTF8Encoding]::new($false)
$template = [IO.File]::ReadAllText($templatePath, $utf8)

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
    [IO.File]::ReadAllText($includePath, $utf8)
})

[IO.File]::WriteAllText($outputPath, $staticHtml, $utf8)
Write-Output "Built $outputPath"