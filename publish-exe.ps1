param(
    [string]$Runtime = "win-x64"
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$project = Join-Path $root "FarmtekTimerRouter.csproj"
$publishDir = Join-Path $root ".publish\$Runtime"
$targetExe = Join-Path $root "FarmtekTimerRouter.exe"
$sourceExe = Join-Path $publishDir "FarmtekTimerRouter.exe"

Push-Location $root
try {
    & dotnet publish $project `
        -c Release `
        -r $Runtime `
        --self-contained true `
        -p:PublishSingleFile=true `
        -p:IncludeNativeLibrariesForSelfExtract=true `
        -p:EnableCompressionInSingleFile=true `
        -o $publishDir

    if ($LASTEXITCODE -ne 0) {
        throw "dotnet publish failed."
    }

    if (-not (Test-Path $sourceExe)) {
        throw "Publish finished but EXE was not found at $sourceExe"
    }

    Copy-Item $sourceExe $targetExe -Force
    Write-Host "EXE ready: $targetExe"
    Write-Host "Publish folder: $publishDir"
} finally {
    Pop-Location
}
