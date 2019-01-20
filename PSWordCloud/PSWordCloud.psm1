$script:ModuleRoot = $PSScriptRoot
Write-Verbose "PSWordCloud module root: $script:ModuleRoot"

# PowerShell Core uses System.Drawing.Common assembly instead of System.Drawing
if ($PSEdition -eq 'Core') {
    Write-Verbose 'Importing necessary types.'
    Add-Type -AssemblyName 'System.Drawing.Common'
}
else {
    Write-Verbose 'Importing necessary types.'
    Add-Type -AssemblyName 'System.Drawing'
}

$PublicFunctions = Get-ChildItem "$script:ModuleRoot\Public"
$PrivateFunctions = Get-ChildItem "$script:ModuleRoot\Private"

foreach ($Function in @($PublicFunctions) + @($PrivateFunctions)) {
    Write-Verbose "Importing functions from file: [$($Function.Name)]"
    . $Function.Fullname
}

Export-ModuleMember -Function $PublicFunctions.BaseName