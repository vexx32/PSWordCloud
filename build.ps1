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

if (Test-Path -Path $OutputPath) {
    Remove-Item -Recurse -Path $OutputPath
}

Start-Process -NoNewWindow -Wait -FilePath 'dotnet' -ArgumentList @(
    'publish'
    "-c $Channel"
    "-o $(Join-Path -Path $OutputPath -ChildPath "bin")"
    $ProjectFile
)

$SupportedPlatforms = "win-x64", "win-x86", "linux-x64", "osx"
$ModulePath = Join-Path -Path $OutputPath "PSWordCloud"
New-Item -Path $ModulePath -ItemType Directory | Out-Null

# Copy the main module DLL to final module directory
Join-Path -Path $OutputPath -ChildPath "bin" "PSWordCloudCmdlet.dll" | Copy-Item -Destination $ModulePath

# Get the main Skia DLL
$SkiaDLL = Join-Path -Path $OutputPath -ChildPath "bin" "SkiaSharp.dll" | Get-Item

# Get platform-specific runtime library folders for Skia
$RuntimeFolders = Join-Path -Path $OutputPath -ChildPath "bin" "runtimes" |
    Get-ChildItem |
    Where-Object Name -in $SupportedPlatforms

$RuntimeFolders | ForEach-Object {
    $OutputDirectory = New-Item -ItemType Directory -Path (Join-Path -Path $ModulePath -ChildPath $_.Name)
    Get-ChildItem -Path $_.FullName -Recurse -Include *.dylib, *.dll, *.so |
        Copy-Item -Destination $OutputDirectory.FullName -PassThru |
        ForEach-Object { Copy-Item -Path $SkiaDLL.FullName -Destination $OutputDirectory.FullName }
}

Split-Path -Path $ProjectPath -Parent |
    Get-ChildItem -Recurse -Include "*.ps*1" |
    Copy-Item -Destination $ModulePath
