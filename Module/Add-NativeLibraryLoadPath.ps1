if ($PSVersionTable.PSVersion.Major -ge 7) {
    return
}

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

Add-Type -TypeDefinition @"
using System;
using System.IO;
using System.Reflection;
using System.Runtime.Loader;

namespace PSWordCloud
{
    public class NativeLoadContext : AssemblyLoadContext
    {
        protected override IntPtr LoadUnmanagedDll(string unmanagedDllName)
        {
            return LoadUnmanagedDllFromPath(Path.Combine($NativeRuntimeFolder, unmanagedDllName));
        }
    }
}
"@
