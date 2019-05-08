switch ($true) {
    $IsMacOS {
        $NativeRuntimeFolder = Join-Path -Path $PSScriptRoot -ChildPath "osx"
        $ConfigPath = Join-Path -Path $NativeRuntimeFolder -ChildPath "libSkiaSharp.dylib.config"
        @"
<configuration>
  <dllmap dll="libSkiaSharp" target="libSkiaSharp.dylib" />
</configuration>
"@ | Set-Content -Path $ConfigPath

        $env:DYLD_FRAMEWORK_PATH = "{0}{1}{2}" -f @(
            $NativeRuntimeFolder
            [IO.Path]::PathSeparator
            $env:DYLD_FRAMEWORK_PATH
        )
    }
    $IsLinux {
        $NativeRuntimeFolder = Join-Path -Path $PSScriptRoot -ChildPath "linux-x64"
        $env:LD_LIBRARY_PATH = "{0}{1}{2}" -f @(
            $NativeRuntimeFolder
            [IO.Path]::PathSeparator
            $env:LD_LIBRARY_PATH
        )
    }
    default {
        # Windows PowerShell / PS Core on Windows
        Add-Type -TypeDefinition @"
            using System.Runtime.InteropServices;

            public class DllLoadPath
            {
                [DllImport("kernel32", CharSet=CharSet.Unicode)]
                public static extern int SetDllDirectory(string NewDirectory);
            }
"@
        $PlatformFolder = if ([Environment]::Is64BitProcess) { "win-x64" } else { "win-x86" }
        $NativeRuntimeFolder = Join-Path -Path $PSScriptRoot -ChildPath $PlatformFolder

        [DllLoadPath]::SetDllDirectory($NativeRuntimeFolder)
    }
}

