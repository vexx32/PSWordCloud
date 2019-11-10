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

$Dotnet = Start-Process -NoNewWindow -PassThru -FilePath 'dotnet' -ArgumentList @(
    'publish'
    "-c $Channel"
    '-o "{0}"' -f (Join-Path -Path $OutputPath -ChildPath "bin")
    $ProjectFile
)

$Dotnet.WaitForExit()

$SupportedPlatforms = "win-x64", "win-x86", "linux-x64", "osx"
$ModulePath = Join-Path $OutputPath -ChildPath "PSWordCloud"
New-Item -Path $ModulePath -ItemType Directory | Out-Null

# Copy the main module DLLs to final module directory
Copy-Item -Path "$OutputPath/bin/PSWordCloudCmdlet.dll" -Destination $ModulePath
Copy-Item -Path "$OutputPath/bin/SkiaSharp.dll" -Destination $ModulePath

# Copy platform-specific runtime library folders for Skia
$RuntimeFolders = Get-ChildItem -Path "$OutputPath/bin/runtimes" |
    Where-Object Name -in $SupportedPlatforms

$RuntimeFolders | ForEach-Object {
    $OutputDirectory = New-Item -ItemType Directory -Path "$ModulePath/$($_.Name)"
    Get-ChildItem -Path $_.FullName -Recurse -Include *.dylib, *.dll, *.so |
        Copy-Item -Destination $OutputDirectory.FullName
}

Split-Path -Path $ProjectFile -Parent |
    Get-ChildItem -Recurse -Include "*.ps*1" |
    Copy-Item -Destination $ModulePath

Import-Module $ModulePath

$DocsPath = Join-Path $PSScriptRoot -ChildPath "docs"
Update-MarkdownHelp -Path $DocsPath -AlphabeticParamsOrder

$ExternalHelpPath = Join-Path $ModulePath -ChildPath "en-US"
New-ExternalHelp -Path $DocsPath -OutputPath $ExternalHelpPath
