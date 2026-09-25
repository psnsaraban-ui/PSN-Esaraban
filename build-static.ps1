$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$outputPath = Join-Path $root 'index.html'
$utf8 = [Text.UTF8Encoding]::new($false)
$appScriptUrl = 'https://script.google.com/macros/s/AKfycbwRY0K59FddGG8F_VdLgHoQZRZImkqyVHXNa9ASzPvoUGx8g1k8AJj-jdOdw48Z2eMQGg/exec'
$redirectHtml = @"
<!doctype html>
<html lang="en">
<head>
    <meta charset="utf-8">
    <meta http-equiv="refresh" content="0;url=$appScriptUrl">
    <title>PSN Saraban</title>
</head>
<body>
    <p>กำลังเปิดระบบ PSN Saraban...</p>
    <script>window.location.replace('$appScriptUrl');</script>
</body>
</html>
"@
[IO.File]::WriteAllText($outputPath, $redirectHtml, $utf8)
Write-Output "Built $outputPath"