[CmdletBinding()]
param(
    [switch]$SkipZip
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$projectPath = Join-Path $repoRoot 'src\AdvancedRdp\AdvancedRdp.csproj'
$publishDir = Join-Path $repoRoot 'dist\advancedrdp-win-x64'
$zipPath = Join-Path $repoRoot 'dist\AdvancedRdp-win-x64-self-contained.zip'

if (Test-Path $publishDir) {
    Remove-Item -Path $publishDir -Recurse -Force
}

Get-ChildItem -Path (Join-Path $repoRoot 'dist') -Filter 'AdvancedRdp*-win-x64*.zip' -ErrorAction SilentlyContinue |
    Remove-Item -Force

if (Test-Path $zipPath) {
    Remove-Item -Path $zipPath -Force
}

dotnet publish $projectPath -p:PublishProfile=SelfContained-win-x64

if (-not (Test-Path (Join-Path $publishDir 'AdvancedRdp.exe'))) {
    throw 'Publish failed: AdvancedRdp.exe was not created.'
}

if (-not $SkipZip) {
    Compress-Archive -Path (Join-Path $publishDir '*') -DestinationPath $zipPath
}

$exe = Get-Item (Join-Path $publishDir 'AdvancedRdp.exe')
$summary = [pscustomobject]@{
    PublishDir = $publishDir
    ExeSizeMB = [Math]::Round($exe.Length / 1MB, 2)
    ZipPath = if ($SkipZip) { '' } else { $zipPath }
}

$summary | Format-List
