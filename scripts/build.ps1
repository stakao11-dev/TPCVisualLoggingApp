<#
.SYNOPSIS
    Visual Studio 2019 の MSBuild でソリューションをビルドする。

.DESCRIPTION
    MSBuild は PATH に無いため、このスクリプトで VS2019 の MSBuild.exe を探して呼び出す。
    Claude Code の許可設定 (.claude/settings.json) ではこのスクリプトだけをビルド用に許可している。
    ビルドのみを行い、生成した exe の実行 (= 実機コントローラとの通信) は行わない。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File scripts/build.ps1
    powershell -ExecutionPolicy Bypass -File scripts/build.ps1 -Configuration Release
    powershell -ExecutionPolicy Bypass -File scripts/build.ps1 -Solution src/Foo.sln -Platform x64
#>
param(
    [string]$Solution,
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Debug',
    [string]$Platform
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot

function Find-MSBuild {
    # 1. vswhere で VS2019 (16.x) を探す
    $vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
    if (Test-Path $vswhere) {
        $found = & $vswhere -version '[16.0,17.0)' -products * -requires Microsoft.Component.MSBuild `
            -find 'MSBuild\**\Bin\MSBuild.exe' | Select-Object -First 1
        if ($found) { return $found }
    }
    # 2. vswhere に登録されていない場合は既定のインストール先を探す
    foreach ($edition in 'Enterprise', 'Professional', 'Community', 'BuildTools') {
        $path = Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\2019\$edition\MSBuild\Current\Bin\MSBuild.exe"
        if (Test-Path $path) { return $path }
    }
    throw 'Visual Studio 2019 の MSBuild.exe が見つかりません。'
}

if (-not $Solution) {
    $slns = @(Get-ChildItem -Path $root -Filter *.sln -Recurse -Depth 2 |
        Where-Object { $_.FullName -notmatch '\\(\.venv|\.cache|node_modules)\\' })
    if ($slns.Count -eq 0) { throw 'ソリューション (.sln) が見つかりません。-Solution で指定してください。' }
    if ($slns.Count -gt 1) { throw "ソリューションが複数あります。-Solution で指定してください:`n$($slns.FullName -join "`n")" }
    $Solution = $slns[0].FullName
}

$msbuild = Find-MSBuild
$msbuildArgs = @($Solution, '-restore', '-m', '-nologo', '-v:minimal', "-p:Configuration=$Configuration")
if ($Platform) { $msbuildArgs += "-p:Platform=$Platform" }

Write-Host "MSBuild : $msbuild"
Write-Host "Solution: $Solution ($Configuration$(if ($Platform) { "|$Platform" }))"
& $msbuild @msbuildArgs
exit $LASTEXITCODE
