[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]
    $Path,

    [Parameter(Mandatory)]
    [string]
    $Name
)

$RepositoryFolder = if (-not (Test-Path $Path)) {
    Write-Host "Folder does not exist; creating '$Path'"
    New-Item -ItemType Directory -Path $Path -Force
}
else {
    Write-Host "Getting existing folder '$Path'"
    Get-Item -Path $Path
}

Write-Host 'Registering Filesystem Repository'
$Params = @{
    Name                 = $Name
    SourceLocation       = $RepositoryFolder.FullName
    ScriptSourceLocation = $RepositoryFolder.FullName
    InstallationPolicy   = 'Trusted'
}
Register-PSRepository @Params
