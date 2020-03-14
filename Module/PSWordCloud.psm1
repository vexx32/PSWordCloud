if ($PSVersionTable.PSVersion.Major -ge 7 -or 'PSWordCloud.PSWordCloudCmdlet' -as [type]) {
    return
}

$PlatformFolder = switch ($true) {
    $IsMacOS { "osx" }
    $IsLinux { "linux-x64" }
    default {
        # Windows
        if ([Environment]::Is64BitProcess) { "win-x64" } else { "win-x86" }
    }
}

$NativeRuntimeFolder = Join-Path -Path $PSScriptRoot -ChildPath $PlatformFolder

Add-Type -TypeDefinition @"
using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Loader;

namespace PSWordCloud
{
    public class LoadContext : AssemblyLoadContext
    {
        protected override Assembly Load(AssemblyName assemblyName)
        {
            if (assemblyName.Name == "PSWordCloudCmdlet")
            {
                return LoadFromAssemblyPath(Path.Combine("$PSScriptRoot","PSWordCloudCmdlet.dll"));
            }

            if (assemblyName.Name == "SkiaSharp")
            {
                return LoadFromAssemblyPath(Path.Combine("$PSScriptRoot","SkiaSharp.dll"));
            }

            // Return null to fallback on default load context
            return null;
        }
        protected override IntPtr LoadUnmanagedDll(string unmanagedDllName)
        {
            if (unmanagedDllName == "liblibSkiaSharp")
            {
                unmanagedDllName = "libSkiaSharp"
            }

            string libExtension = "dll";
            if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
            {
                libExtension = "dylib";
            }
            elseif (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
            {
                libExtension = "so";
            }

            return LoadUnmanagedDllFromPath(Path.Combine("$NativeRuntimeFolder", $"unmanagedDllName.{libExtension}"));
        }
    }
}
"@

$Context = [PSWordCloud.LoadContext]::new()
$Context.Load([System.Reflection.AssemblyName]::new("SkiaSharp"))
$Context.Load([System.Reflection.AssemblyName]::new("PSWordCloudCmdlet"))
