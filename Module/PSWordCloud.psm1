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
                unmanagedDllName = "libSkiaSharp";
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
