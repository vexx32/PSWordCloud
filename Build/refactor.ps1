[cmdletbinding()]
param (
    [Parameter(Mandatory)]
    [string]$PSD1Path,

    [Parameter()]
    [string[]]$CopyDirectories,

    [Parameter()]
    [switch]$OverwritePSM1
)

begin {
    $eap = $ErrorActionPreference
    $ErrorActionPreference = 'Stop'     # Make errors terminate instead of fighting through them
}

process {
    $moduleBase = Split-Path $PSD1Path -Parent
    $moduleName = (Split-Path $PSD1Path -Leaf) -replace '.psd1$'
    $outputDir = "$moduleBase\..\Output"

    try {
        Write-Verbose "Creating output directory"
        $null = Remove-Item -Path $outputDir -Recurse -Force -ErrorAction SilentlyContinue
        $null = New-Item -Path $outputDir -ItemType Directory
    }
    catch {
        throw "Unable to create output directory: $outputDir"
    }

    try {
        Write-Verbose "Copying psm1, psd1 and ps1xml files to output"
        $null = Copy-Item "$moduleBase\*.psm1", "$moduleBase\*.psd1", "$moduleBase\*.ps1xml" -Destination $outputDir -Force
    }
    catch {
        throw "Unable to gather the root files"
    }

    if ($CopyDirectories) {
        Write-Verbose "Copying additonal items to output"
        try {
            $null = Copy-Item -Path $CopyDirectories -Recurse -Destination $outputDir -Force
        }
        catch {
            throw "Unable to gather copy directories"
        }
    }

    #try {
        Write-Verbose "Contactenating public and private functions into PSM1 in output dir"
        $public = Get-ChildItem -Path "$($moduleBase)\Public\*.ps1"
        $private = Get-ChildItem -Path "$($moduleBase)\Private\*.ps1"

        $psm1Content = @(@($public) + @($private)) | foreach-object {
            Get-Content $_.FullName
        }
        
        if ($OverwritePSM1) {
            $psm1Content | Set-Content -Path "$outputDir\$($moduleName).psm1" 
        }
        else {
            $psm1Content | Add-Content -Path "$outputDir\$($moduleName).psm1" 
        }
    #}
    #catch {
    #    throw "Unable to combine PS1 files into PSM1"
    #}

    #try {
        Write-Verbose "Updating FunctionsToExport property in output PSD1"
        Update-ModuleManifest -Path  "$outputDir\$moduleName.psd1" -FunctionsToExport $public.BaseName
    #}
    #catch {
    #    throw "Unable to update the FunctionsToExport property in output PSD1"
    #}
}

end {
    $ErrorActionPreference = $eap
}