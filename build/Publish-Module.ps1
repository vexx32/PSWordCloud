[CmdletBinding()]
param(
    [string]
    $Key,

    [string]
    $Path,

    [string]
    $OutputDirectory,

    [string]
    $ModulePath
)

$env:NugetApiKey = $Key

$ModuleName = Split-Path $ModulePath -Leaf

if ($OutputDirectory) {
    Import-Module $ModulePath
    $Module = Get-Module -Name $ModuleName
    $Dependencies = @(
        $Module.RequiredModules.Name
        $Module.NestedModules.Name
    ).Where{ $_ }

    foreach ($Module in $Dependencies) {
        Publish-Module -Name $Module -Repository FileSystem -NugetApiKey "Test-Publish"
    }
}

$HelpFile = Get-ChildItem -Path "$ModulePath" -File -Recurse -Filter '*-help.xml'

if ($HelpFile.Directory -notmatch 'en-us|\w{1,2}-\w{1,2}') {
    $PSCmdlet.WriteError(
        [System.Management.Automation.ErrorRecord]::new(
            [IO.FileNotFoundException]::new("Help files are missing!"),
            'Build.HelpXmlMissing',
            'ObjectNotFound',
            $null
        )
    )

    exit 404
}

$DeploymentParams = @{
    Path    = $Path
    Recurse = $false
    Force   = $true
    Verbose = $true
}

Invoke-PSDeploy @DeploymentParams

Get-ChildItem -Path $DeploymentParams['Path'] | Out-String | Write-Host

$Nupkg = Get-ChildItem -Path $DeploymentParams['Path'] -Filter "$ModuleName*.nupkg" | ForEach-Object FullName
Write-Host "##vso[task.setvariable variable=NupkgPath]$Nupkg"
