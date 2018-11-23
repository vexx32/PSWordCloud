$script:ModuleRoot = $PSScriptRoot
Write-Verbose "PSWordCloud module root: $script:ModuleRoot"

$PublicFunctions = Get-ChildItem "$script:ModuleRoot\Public"
$PrivateFunctions = Get-ChildItem "$script:ModuleRoot\Private"

foreach ($Function in @($PublicFunctions) + @($PrivateFunctions)) {
    Write-Verbose "Importing functions from file: [$($Function.Name)]"
    . $Function.Fullname
}

# PowerShell Core uses System.Drawing.Common assembly instead of System.Drawing
if ($PSEdition -eq 'Core') {
    Write-Verbose 'Importing necessary types.'
    Add-Type -AssemblyName 'System.Drawing.Common'
}
else {
    Write-Verbose 'Importing necessary types.'
    Add-Type -AssemblyName 'System.Drawing'
}

Export-ModuleMember -Function $PublicFunctions.BaseName