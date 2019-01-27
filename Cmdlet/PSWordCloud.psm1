# Add library folders to necessary path vars
Get-ChildItem -Directory -Path "$PSScriptRoot\bin\release\netstandard2.0\publish\runtimes\" -Recurse -Filter 'native' |
    ForEach-Object {
    $Path = $_
    switch ($true) {
        ($IsWindows -or $PSVersionTable.PSVersion.Major -eq 5) {
            $env:PATH = '{0}{1}{2}' -f @(
                $Path.FullName
                [System.IO.Path]::PathSeparator
                $env:PATH
            )
        }
        $IsLinux {
            $env:LD_LIBRARY_PATH = '{0}{1}{2}' -f @(
                $Path.FullName
                [System.IO.Path]::PathSeparator
                $env:LD_LIBRARY_PATH
            )
        }
        $IsMacOS {
            $env:DYLD_LIBRARY_PATH = '{0}{1}{2}' -f @(
                $Path.FullName
                [System.IO.Path]::PathSeparator
                $env:DYLD_LIBRARY_PATH
            )
        }
    }
}

Add-Type -Path "$PSScriptRoot\bin\debug\netstandard2.0\publish\SkiaSharp.dll"
Import-Module "$PSScriptRoot\bin\debug\netstandard2.0\publish\PSWordCloudCmdlet.dll"

Export-ModuleMember -Cmdlet 'New-WordCloud'