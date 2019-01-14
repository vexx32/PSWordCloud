using namespace System.Collections.Generic
using namespace System.Drawing
using namespace System.IO
using namespace System.Management.Automation
using namespace System.Numerics

$script:ModuleRoot = $PSScriptRoot

# PowerShell Core uses System.Drawing.Common assembly instead of System.Drawing
if ($PSEdition -eq 'Core') {
    Write-Verbose 'Importing necessary types.'
    Add-Type -AssemblyName 'System.Drawing.Common'
}
else {
    Write-Verbose 'Importing necessary types.'
    Add-Type -AssemblyName 'System.Drawing'
}