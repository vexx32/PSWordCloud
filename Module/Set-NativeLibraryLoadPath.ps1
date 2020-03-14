switch ($PSVersionTable.PSVersion.Major) {
    { $_ -ge 7 } {
        return
    }
    6 {
        if (-not $IsWindows) {
            throw "Cannot load native dependencies in PowerShell 6.x on Unix systems; please install PowerShell 7 or higher."
        }
    }
}

if ( 'PSWordCloud.PSWordCloudCmdlet' -as [type] ) {
    return
}

Add-Type -TypeDefinition @"
    using System.Runtime.InteropServices;
    public class DllLoadPath
    {
        [DllImport("kernel32", CharSet=CharSet.Unicode)]
        private static extern int SetDllDirectory(string NewDirectory);

        public static void SetLoadPath(string path)
        {
            SetDllDirectory(path);
        }
    }
"@

$PlatformFolder = if ([Environment]::Is64BitProcess) { "win-x64" } else { "win-x86" }

$NativeRuntimeFolder = Join-Path -Path $PSScriptRoot -ChildPath $PlatformFolder
[DllLoadPath]::SetLoadPath($NativeRuntimeFolder)
