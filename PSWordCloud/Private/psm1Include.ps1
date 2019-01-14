# PowerShell Core uses System.Drawing.Common assembly instead of System.Drawing
if ($PSEdition -eq 'Core') {
    Write-Verbose 'Importing necessary types.'
    Add-Type -AssemblyName 'System.Drawing.Common'
}
else {
    Write-Verbose 'Importing necessary types.'
    Add-Type -AssemblyName 'System.Drawing'
}