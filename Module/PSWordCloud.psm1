Add-Type -TypeDefinition @"
    using System.Runtime.InteropServices;

    public class DllLoadPath
    {
        [DllImport("kernel32", CharSet=CharSet.Unicode)]
        public static extern int AddDllDirectory(string NewDirectory);
    }
"@

$PlatformFolder = switch ($true) {
    $IsWindows {
        if ([Environment]::Is64BitProcess) { "win-x64" } else { "win-x86" }
    }
    $IsMacOS { "osx" }
    $IsLinux { "linux-x64" }
    default {
        # Windows PowerShell
        if ([Environment]::Is64BitProcess) { "win-x64" } else { "win-x86" }
    }
}

$NativeRuntimeFolder = Join-Path -Path $PSScriptRoot -ChildPath $PlatformFolder
[DllLoadPath]::AddDllDirectory($NativeRuntimeFolder)

$SkiaDllPath = Join-Path -Path $PSScriptRoot -ChildPath "SkiaSharp.dll"
Add-Type -Path $SkiaDllPath

$ModuleDllPath = Join-Path -Path $PSScriptRoot -ChildPath "PSWordCloudCmdlet.dll"
Import-Module $ModuleDllPath
