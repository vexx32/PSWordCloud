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

foreach ($RuntimeId in 'win-x64', 'linux-x64', 'osx-x64') {
    $args = @(
        'publish'
        '-c', $Channel
        '-o', ('"{0}"' -f (Join-Path -Path $OutputPath -ChildPath "bin"))
        '-r', $RuntimeId
        '--self-contained', 'true'
        '"{0}"' -f (Get-Item $ProjectFile).FullName
    )

    & dotnet @args
}

$ModulePath = Join-Path $OutputPath -ChildPath "PSWordCloud"
New-Item -Path $ModulePath -ItemType Directory | Out-Null

# Copy the main module DLLs to final module directory
Copy-Item -Path "$OutputPath/bin/PSWordCloudCmdlet.dll" -Destination $ModulePath
Copy-Item -Path "$OutputPath/bin/SkiaSharp.dll" -Destination $ModulePath

# Copy native runtimes
$RuntimeFolders = Get-Item -Path "$OutputPath/bin/libSkiaSharp.*" |
    ForEach-Object {
        $File = $_
        switch ($_.Extension) {
            '.dll' {
                $Destination = New-Item -ItemType Directory -Path "$ModulePath/win-x64"
                $File | Copy-Item -Destination $Destination.Fullname
            }
            '.so' {
                $Destination = New-Item -ItemType Directory -Path "$ModulePath/linux-x64"
                $File | Copy-Item -Destination $Destination.Fullname
            }
            '.dylib' {
                $Destination = New-Item -ItemType Directory -Path "$ModulePath/osx-x64"
                $File | Copy-Item -Destination $Destination.Fullname
            }
        }
    }

Split-Path -Path $ProjectFile -Parent |
    Get-ChildItem -Recurse -Include "*.ps*1" |
    Copy-Item -Destination $ModulePath

Import-Module $ModulePath

$DocsPath = Join-Path $PSScriptRoot -ChildPath "docs"
Update-MarkdownHelp -Path $DocsPath -AlphabeticParamsOrder

$ExternalHelpPath = Join-Path $ModulePath -ChildPath "en-US"
New-ExternalHelp -Path $DocsPath -OutputPath $ExternalHelpPath
