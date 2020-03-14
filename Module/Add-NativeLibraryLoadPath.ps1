if ($PSVersionTable.PSVersion.Major -ge 7) {
    return
}

Add-Type -TypeDefinition @"
using System;
using System.IO;
using System.Runtime.InteropServices;

namespace PSWordCloud
{
    [Flags]
    internal enum DLOpenFlags
    {
        RTLD_LAZY = 1,
        RTLD_NOW = 2,
        RTLD_LOCAL = 4,
        RTLD_GLOBAL = 8,
    }

    internal static class UnixNativeMethods
    {
        [DllImport("libdl", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr dlopen(
            string dlToOpen,
            DLOpenFlags flags);
    }

    internal static class Win32NativeMethods
    {
        [DllImport("kernel32", EntryPoint="LoadLibraryW")]
        public static extern IntPtr LoadLibrary(
            [InAttribute, MarshalAs(UnmanagedType.LPWStr)] string dllToLoad);
    }

    public static class NativeMethods
    {
        public static IntPtr LibHandle { get; private set; }

        public static void LoadNativeLibrary(string dllPath)
        {
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                LibHandle = Win32NativeMethods.LoadLibrary(dllPath);
            }
            else
            {
                LibHandle = UnixNativeMethods.dlopen(dllPath, DLOpenFlags.RTLD_NOW);
            }
        }
    }
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
$DllPath = Get-Item -Path "$NativeRuntimeFolder/libSkiaSharp.*"

[PSWordCloud.NativeMethods]::LoadNativeLibrary($DllPath.FullName)
