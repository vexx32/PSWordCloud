Add-Type -ReferencedAssemblies 'System.Runtime.Loader' -TypeDefinition @"
    using System;
    using System.Reflection;
    using System.Runtime.Loader;

    namespace PSWordCloud.Native
    {
        public class Loader : AssemblyLoadContext
        {
            private static Loader singleton = new Loader();
            private Loader() : base() {}

            public static IntPtr LoadLibrary(string path)
            {
                return singleton.LoadUnmanagedDllFromPath(path);
            }

            protected override Assembly Load(AssemblyName assemblyName) => null;
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
