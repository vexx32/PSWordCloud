Add-Type -ReferencedAssemblies 'System.Runtime.Loader' -TypeDefinition @"
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
            static void LoadLibrary(string path)
            {
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                {
                    Win32NativeMethods.LoadLibrary(path);
                }
                else
                {
                    LinuxNativeMethods.dlopen(path, DLOpenFlags.RTLD_LAZY);
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

$NativeRuntime = Get-Item -Path "$PSScriptRoot/$PlatformFolder/libSkiaSharp.*"
[PSWordCloud.Native.Loader]::LoadLibrary($NativeRuntime)
