$PlatformFolder = switch ($true) {
    $IsWindows {
        if ([Environment]::Is64BitProcess) { "win-x64" } else { "win-x86" }
    }
    $IsMacOS {
        "osx"
    }
    $IsLinux {
        "linux-x64"
    }
    default {
        # Windows PowerShell
        if ([Environment]::Is64BitProcess) { "win-x64" } else { "win-x86" }
    }
}

$SkiaDllPath = Join-Path -Path $PSScriptRoot -ChildPath $PlatformFolder |
    Join-Path -ChildPath "SkiaSharp.dll"

Add-Type -Path $SkiaDllPath

$ModuleDllPath = Join-Path -Path $PSScriptRoot -ChildPath "PSWordCloudCmdlet.dll"
Import-Module $ModuleDllPath

Export-ModuleMember -Cmdlet "New-WordCloud"
