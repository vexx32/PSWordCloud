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
$ModulePath = "$OutputPath/PSWordCloud"
New-Item -Path $ModulePath -ItemType Directory | Out-Null

# Copy the main module DLL to final module directory
Copy-Item -Path "$OutputPath/bin/PSWordCloudCmdlet.dll" -Destination $ModulePath

# Get the main Skia DLL
$SkiaDLL = Get-Item -Path "$OutputPath/bin/SkiaSharp.dll"

# Get platform-specific runtime library folders for Skia
$RuntimeFolders = Get-ChildItem -Path "$OutputPath/bin/runtimes" |
    Where-Object Name -in $SupportedPlatforms

$RuntimeFolders | ForEach-Object {
    $OutputDirectory = New-Item -ItemType Directory -Path "$ModulePath/$($_.Name)"
    Get-ChildItem -Path $_.FullName -Recurse -Include *.dylib, *.dll, *.so |
        Copy-Item -Destination $OutputDirectory.FullName -PassThru |
        ForEach-Object { Copy-Item -Path $SkiaDLL.FullName -Destination $OutputDirectory.FullName }
}

Split-Path -Path $ProjectFile -Parent |
    Get-ChildItem -Recurse -Include "*.ps*1" |
    Copy-Item -Destination $ModulePath
