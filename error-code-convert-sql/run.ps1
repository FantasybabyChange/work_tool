param(
    [Parameter(Mandatory=$true)]
    [string]$ExcelPath,
    [string]$OutputPath,
    [string]$Sheet,
    [string]$Prefix
)

$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$JarPath = Join-Path $Root "target\error-code-convert-sql-1.0.0.jar"

if (-not (Test-Path $JarPath)) {
    Write-Host "JAR 不存在，正在编译..."
    & (Join-Path $Root "build.ps1")
    if ($LASTEXITCODE -ne 0) {
        throw "编译失败"
    }
}

$javaArgs = @("-jar", $JarPath, $ExcelPath)

if ($OutputPath) { $javaArgs += "-o"; $javaArgs += $OutputPath }
if ($Sheet) { $javaArgs += "--sheet"; $javaArgs += $Sheet }
if ($Prefix) { $javaArgs += "--prefix"; $javaArgs += $Prefix }

& java @javaArgs
