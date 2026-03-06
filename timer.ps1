param(
    [switch]$NoBrowser,
    [switch]$ValidateOnly,
    [int]$RunForSeconds = 0,
    [int]$UiPort = 0
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$project = Join-Path $root "FarmtekTimerRouter.csproj"
$dll = Join-Path $root "bin\\Release\\net8.0-windows\\FarmtekTimerRouter.dll"
$sources = Get-ChildItem -Path $root -Filter *.cs -File | Select-Object -ExpandProperty FullName

if (-not (Test-Path $project)) {
    throw "Missing project file: $project"
}

$needsBuild = -not (Test-Path $dll)
if (-not $needsBuild) {
    $dllWriteTime = (Get-Item $dll).LastWriteTimeUtc
    foreach ($path in @($project) + $sources) {
        if ((Get-Item $path).LastWriteTimeUtc -gt $dllWriteTime) {
            $needsBuild = $true
            break
        }
    }
}

if ($needsBuild) {
    Push-Location $root
    try {
        & dotnet build $project -c Release --nologo | Out-Host
    } finally {
        Pop-Location
    }
}

$arguments = @($dll)

if ($NoBrowser) {
    $arguments += "--no-browser"
}

if ($ValidateOnly) {
    $arguments += "--validate-only"
}

if ($RunForSeconds -gt 0) {
    $arguments += "--run-for-seconds"
    $arguments += $RunForSeconds.ToString()
}

if ($UiPort -gt 0) {
    $arguments += "--port"
    $arguments += $UiPort.ToString()
}

Push-Location $root
try {
    & dotnet @arguments
} finally {
    Pop-Location
}
