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

$PSVersionTable | Out-String | Write-Host

Write-Host "Importing PlatyPS"
Import-Module PlatyPS

if (Test-Path -Path $OutputPath) {
    Write-Host "Cleaning up '$OutputPath'"
    Get-ChildItem -Directory $OutputPath -Include 'bin', 'PSWordCloud' |
        Remove-Item -Recurse
}

$ModulePath = Join-Path $OutputPath -ChildPath 'PSWordCloud'

$SupportedPlatforms = "win-x64", "win-x86", "linux-x64", "osx"

foreach ($rid in $SupportedPlatforms) {
    $binPath = Join-Path -Path $OutputPath -ChildPath "bin"
    if (Test-Path $binPath) {
        Remove-Item -Recurse -Path $binPath -Force
    }

    Write-Host "Running 'dotnet publish' with RID: $rid"
    $process = Start-Process -FilePath 'dotnet' -WindowStyle Hidden -PassThru -ArgumentList @(
        'publish'
        "--configuration $Channel"
        "--output $binPath"
        $ProjectFile
        "--runtime $rid"
    )

    try {
        Write-Host "Waiting for dotnet process to exit..."
        $process.WaitForExit()
    }
    catch {
        Write-Warning "dotnet process errored on WaitForExit()"
    }

    Write-Host "Locating 'libSkiaSharp' file for $rid"
    $nativeLib = Join-Path $OutputPath -ChildPath 'bin' |
        Get-ChildItem -Recurse -File -Filter '*libSkiaSharp*'
    $destinationPath = Join-Path $ModulePath -ChildPath $rid |
        New-Item -Path { $_ } -ItemType Directory |
        Select-Object -ExpandProperty FullName

    Write-Host "Moving $nativeLib to $destinationPath"
    $nativeLib = $nativeLib | Move-Item -Destination $destinationPath -Force -PassThru

    if ($rid -notmatch '^win') {
        $nativeLib | Rename-Item -NewName { "SkiaSharp$($_.Extension)" }
    }
}

Write-Host "Copying PSWordCloud and SkiaSharp DLLs to '$ModulePath'"
Copy-Item -Path "$OutputPath/bin/PSWordCloudCmdlet.dll" -Destination $ModulePath
Copy-Item -Path "$OutputPath/bin/SkiaSharp.dll" -Destination $ModulePath

Write-Host "Copying PowerShell files to module folder"
Split-Path -Path $ProjectFile -Parent |
    Get-ChildItem -Recurse -Include "*.ps*1" |
    Copy-Item -Destination $ModulePath

Write-Host "Importing the built module"
Import-Module $ModulePath

$DocsPath = Join-Path "$PSScriptRoot/.." -ChildPath "docs"
Write-Host "Updating markdown help files in '$DocsPath'"
Update-MarkdownHelp -Path $DocsPath -AlphabeticParamsOrder

$ExternalHelpPath = Join-Path $ModulePath -ChildPath "en-US"
Write-Host "Generating external help in '$ExternalHelpPath'"
New-ExternalHelp -Path $DocsPath -OutputPath $ExternalHelpPath
