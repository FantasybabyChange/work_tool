$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Classes = Join-Path $Root "target\classes"
$JarPath = Join-Path $Root "target\error-code-convert-sql-1.0.0.jar"
$ManifestPath = Join-Path $Root "target\MANIFEST.MF"
$SourcesPath = Join-Path $Root "target\sources.txt"

New-Item -ItemType Directory -Force -Path $Classes | Out-Null
Get-ChildItem -Path (Join-Path $Root "src\main\java") -Recurse -Filter *.java |
    ForEach-Object { $_.FullName } |
    Set-Content -Path $SourcesPath -Encoding ASCII

javac -encoding UTF-8 -d $Classes "@$SourcesPath"
if ($LASTEXITCODE -ne 0) {
    throw "javac failed with exit code $LASTEXITCODE"
}

"Manifest-Version: 1.0`nMain-Class: com.local.tools.errorcode.ErrorCodeConvertSqlApplication`n" |
    Set-Content -Path $ManifestPath -Encoding ASCII

Push-Location $Classes
try {
    jar cfm $JarPath $ManifestPath .
    if ($LASTEXITCODE -ne 0) {
        throw "jar failed with exit code $LASTEXITCODE"
    }
} finally {
    Pop-Location
}

Write-Host "Built $JarPath"
