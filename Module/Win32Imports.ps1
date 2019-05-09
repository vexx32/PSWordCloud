if (-not (Get-Variable -Name IsWindows)) {
    $IsWindows = -not ($IsLinux -or $IsMacOS)
}

if ($IsWindows) {
    Add-Type -TypeDefinition @"
        using System.Runtime.InteropServices;
        public class DllLoadPath
        {
            [DllImport("kernel32", CharSet=CharSet.Unicode)]
            public static extern int SetDllDirectory(string NewDirectory);
        }
"@
    $PlatformFolder = if ([Environment]::Is64BitProcess) { 'win-x86' } else { 'win-x64' }
    $NativeRuntimeFolder = Join-Path -Path $PSScriptRoot -ChildPath $PlatformFolder
    [DllLoadPath]::SetDllDirectory($NativeRuntimeFolder)
}
