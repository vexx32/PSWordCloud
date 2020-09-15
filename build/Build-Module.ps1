#Requires -Version 6.1
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

$Line = '-' * 50

$PSVersionTable | Out-String | Write-Host

Write-Host $Line
Write-Host "Setup"
Write-Host $Line

Write-Host "Importing PlatyPS"
Import-Module PlatyPS

if (Test-Path -Path $OutputPath) {
    Write-Host "Cleaning up '$OutputPath'"

    Get-ChildItem -Path $OutputPath -Directory |
        Where-Object Name -in 'bin', 'PSWordCloud' |
        Remove-Item -Recurse
}

$ModulePath = Join-Path $OutputPath -ChildPath 'PSWordCloud'

$SupportedPlatforms = "win-x64", "win-x86", "linux-x64", "osx"

Write-Host $Line
Write-Host "Compiling module with dotnet publish"
Write-Host $Line
foreach ($rid in $SupportedPlatforms) {
    $binPath = Join-Path -Path $OutputPath -ChildPath "bin"
    if (Test-Path $binPath) {
        Remove-Item -Recurse -Path $binPath -Force
    }

    Write-Host "Running 'dotnet publish' with RID: $rid"
    $process = Start-Process -FilePath 'dotnet' -NoNewWindow -PassThru -ArgumentList @(
        'publish'
        "--configuration $Channel"
        "--output $binPath"
        $ProjectFile
        "--runtime $rid"
    )

    $process.WaitForExit()

    Write-Host "Locating 'libSkiaSharp' file for $rid"

    $nativeLib = Join-Path $OutputPath -ChildPath 'bin' |
        Get-ChildItem -Recurse -File -Filter '*libSkiaSharp*'

    <#
        SkiaSharp designates the 'osx' RID, but pwsh only recognises the
        'osx-64' RID when looking for native library folders in the module
        directory.
    #>
    if ($rid -eq 'osx') {
        $rid = 'osx-x64'
    }

    $destinationPath = Join-Path $ModulePath -ChildPath $rid |
        New-Item -Path { $_ } -ItemType Directory |
        Select-Object -ExpandProperty FullName

    Write-Host "Moving $nativeLib to $destinationPath"
    $nativeLib = $nativeLib | Copy-Item -Destination $destinationPath -Force -PassThru

    $nativeLib
}

Write-Host $Line
Write-Host "Copying PSWordCloud and SkiaSharp Files to '$ModulePath'"
Write-Host $Line

Copy-Item -Path "$OutputPath/bin/PSWordCloudCmdlet.dll" -Destination $ModulePath
Copy-Item -Path "$OutputPath/bin/SkiaSharp.dll" -Destination $ModulePath

Write-Host "Copying script & metadata files to module folder"

Split-Path -Path $ProjectFile -Parent |
    Get-ChildItem -Recurse -Include "*.ps*1" |
    Copy-Item -Destination $ModulePath

Write-Host $Line
Write-Host "Generating External Help"
Write-Host $Line

Write-Host "Importing the built module"
Import-Module $ModulePath

$DocsPath = Join-Path "$PSScriptRoot/.." -ChildPath "docs"

Write-Host "Updating markdown help files in '$DocsPath'"

Update-MarkdownHelp -Path $DocsPath -AlphabeticParamsOrder

$ExternalHelpPath = Join-Path $ModulePath -ChildPath "en-US"

Write-Host "Generating external help in '$ExternalHelpPath'"

New-ExternalHelp -Path $DocsPath -OutputPath $ExternalHelpPath
