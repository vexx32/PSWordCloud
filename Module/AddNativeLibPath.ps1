﻿$Assemblies = @(
    'System.Runtime.Loader'
    'System.Runtime.InteropServices.RuntimeInformation'
)
Add-Type -ReferencedAssemblies $Assemblies -TypeDefinition @"
    using System;
    using System.Reflection;
    using System.Runtime.InteropServices;
    using System.Runtime.Loader;

    namespace PSWordCloud.Native
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
            public static extern IntPtr dlopen(string dlToOpen, DLOpenFlags flags);
        }

        internal static class Win32NativeMethods
        {
            [DllImport("kernel32.dll", EntryPoint="LoadLibraryW")]
            public static extern IntPtr LoadLibrary([InAttribute, MarshalAs(UnmanagedType.LPWStr)] string dllToLoad);
        }

        public static class Loader
        {
            public static void LoadLibrary(string path)
            {
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                {
                    Win32NativeMethods.LoadLibrary(path);
                }
                else
                {
                    UnixNativeMethods.dlopen(path, DLOpenFlags.RTLD_LAZY);
                }
            }
        }
    }
"@

$PlatformFolder = switch ($true) {
    $IsWindows { "win-x64" }
    $IsMacOS { "osx-x64" }
    $IsLinux { "linux-x64" }
    default {
        # Windows PowerShell
        "win-x64"
    }
}

$NativeRuntime = Get-Item -Path "$PSScriptRoot/$PlatformFolder/libSkiaSharp.*"
[PSWordCloud.Native.Loader]::LoadLibrary($NativeRuntime)
