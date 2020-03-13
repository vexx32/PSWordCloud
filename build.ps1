#Requires -Version 6.1
[CmdletBinding()]
param(
    [Parameter()]
    [ValidateSet("Debug", "Release")]
    [string]
    $Channel = "Debug",

    [Parameter()]
    [string]
    $OutputPath = (Join-Path -Path $PSScriptRoot -ChildPath "build"),

    [Parameter()]
    [string]
    $ProjectFile = (Join-Path -Path $PSScriptRoot -ChildPath "Module" "PSWordCloudCmdlet.csproj")
)

Install-Module -Name PlatyPS -Scope CurrentUser
Import-Module PlatyPS

if (Test-Path -Path $OutputPath) {
    Remove-Item -Recurse -Path $OutputPath
}

$ModulePath = Join-Path $OutputPath -ChildPath 'PSWordCloud'

$SupportedPlatforms = "win-x64", "win-x86", "linux-x64", "osx"

foreach ($rid in $SupportedPlatforms) {
    $binPath = Join-Path -Path $OutputPath -ChildPath "bin"
    $Dotnet = Start-Process -NoNewWindow -PassThru -FilePath 'dotnet' -ArgumentList @(
        'publish'
        "-c $Channel"
        "-o $binPath"
        $ProjectFile
        "-r $rid"
    )
    $Dotnet | Wait-Process

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
