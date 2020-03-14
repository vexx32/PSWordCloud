if ($PSVersionTable.PSVersion.Major -lt 7) {
    Add-Type -TypeDefinition @"
        using System.Runtime.InteropServices;
        public class DllLoadPath
        {
            [DllImport("kernel32", CharSet=CharSet.Unicode)]
            public static extern int SetDllDirectory(string NewDirectory);
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
    [DllLoadPath]::SetDllDirectory($NativeRuntimeFolder)
}
