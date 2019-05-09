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
    $PlatformFolder = if ([Environment]::Is64BitProcess) { 'win-x64' } else { 'win-x86' }
    $NativeRuntimeFolder = Join-Path -Path $PSScriptRoot -ChildPath $PlatformFolder
    [DllLoadPath]::SetDllDirectory($NativeRuntimeFolder)
}
