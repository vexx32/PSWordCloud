﻿#Requires -Version 6.1
[CmdletBinding()]
param(
    [Parameter()]
    [ValidateSet("Debug", "Release")]
    [string]
    $Channel = "Debug",

    [Parameter()]
    [string]
    $OutputPath = $PSScriptRoot,

    [Parameter()]
    [string]
    $ProjectFile = (Join-Path -Path "$PSScriptRoot/.." -ChildPath "Module" "PSWordCloudCmdlet.csproj")
)

Import-Module PlatyPS

if (Test-Path -Path $OutputPath) {
    Remove-Item -Recurse -Path $OutputPath
}

$ModulePath = Join-Path $OutputPath -ChildPath 'PSWordCloud'

$SupportedPlatforms = "win-x64", "win-x86", "linux-x64", "osx"

foreach ($rid in $SupportedPlatforms) {
    $binPath = Join-Path -Path $OutputPath -ChildPath "bin"
    if (Test-Path $binPath) {
        Remove-Item -Recurse -Path $binPath -Force
    }

    $process = Start-Process -FilePath 'dotnet' -PassThru -ArgumentList @(
        'publish'
        "--configuration $Channel"
        "--output $binPath"
        $ProjectFile
        "--runtime $rid"
    )

    try {
        Start-Sleep -Seconds 60
        $process.WaitForExit()
    }
    catch {
        Write-Warning "dotnet process errored on WaitForExit()"
        Start-Sleep -Seconds 10
    }

    if ($LASTEXITCODE -eq 1) { $global:LASTEXITCODE = 0 }

    $nativeLib = Join-Path $OutputPath -ChildPath 'bin' |
        Get-ChildItem -Recurse -File -Filter '*libSkiaSharp*'
    $destinationPath = Join-Path $ModulePath -ChildPath $rid |
        New-Item -Path { $_ } -ItemType Directory |
        Select-Object -ExpandProperty FullName

    $nativeLib | Move-Item -Destination $destinationPath -Force
}

# Copy the main module DLLs to final module directory
Copy-Item -Path "$OutputPath/bin/PSWordCloudCmdlet.dll" -Destination $ModulePath
Copy-Item -Path "$OutputPath/bin/SkiaSharp.dll" -Destination $ModulePath

Split-Path -Path $ProjectFile -Parent |
    Get-ChildItem -Recurse -Include "*.ps*1" |
    Copy-Item -Destination $ModulePath

Import-Module $ModulePath

$DocsPath = Join-Path $PSScriptRoot -ChildPath "docs"
Update-MarkdownHelp -Path $DocsPath -AlphabeticParamsOrder

$ExternalHelpPath = Join-Path $ModulePath -ChildPath "en-US"
New-ExternalHelp -Path $DocsPath -OutputPath $ExternalHelpPath

$host.SetShouldExit(0)
$global:LASTEXITCODE = 0
