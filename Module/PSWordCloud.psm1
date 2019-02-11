$PlatformFolder = switch ($true) {
    $IsWindows {
        if ([IntPtr]::Size -eq 8) {
            "win-x64"
        }
        else {
            "win-x86"
        }
    }
    $IsMacOS {
        "osx"
    }
    $IsLinux {
        "linux-x64"
    }
}

$SkiaDllPath = Join-Path -Path $PSScriptRoot -ChildPath $PlatformFolder "SkiaSharp.dll"

Add-Type -Path $SkiaDllPath
Import-Module  "$PSScriptRoot\PSWordCloudCmdlet.dll"

Export-ModuleMember -Cmdlet "New-WordCloud"