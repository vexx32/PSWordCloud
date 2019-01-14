[cmdletbinding()]
param (
    [Parameter(Mandatory)]
    [string]$PSD1Path,

    [Parameter(Mandatory)]
    [string]$SettingsFile,

    [Parameter()]
    [string]$OutputXML = "$($env:AGENT_WORKFOLDER)\PssaResults.xml"
)

begin {
    try {
        $null = Get-PackageProvider -Name Nuget
    }
    catch {
        $null = Install-PackageProvider -Name Nuget -Scope CurrentUser -Force -Confirm:$false
    }

    if (@(Get-Module -ListAvailable Pester)[0].Version -lt [version]'4.4.0') {
        $null = Install-Module -Name Pester -Scope CurrentUser -Force -Confirm:$false -SkipPublisherCheck
    }

    if (@(Get-Module -ListAvailable PSScriptAnalyzer)[0].Version -lt [version]'1.17.1') {
        $null = Install-Module -Name PSScriptAnalyzer -Scope CurrentUser -Force -Confirm:$false
    }

    Import-Module Pester -Version 4.4.0
    Import-Module PSScriptAnalyzer -Version 1.17.1
}

process {
    $moduleBase = Split-Path $PSD1Path -Parent
    $moduleName = (Split-Path $PSD1Path -Leaf) -replace '.psd1$'


    $testFile = New-TemporaryFile
    Rename-Item -Path $testFile -NewName "$($testFile.BaseName).ps1"
    $testFilePS1 = "$($testFile.Directory)\$($testFile.basename).ps1"

    $testFileContent = @"
    describe 'PSSA' {
        `$allFiles = Get-ChildItem -Path "$($moduleBase)\*.ps1", "$moduleBase\*.psm1", "$moduleBase\*.psd1" -Recurse
        foreach (`$file in `$allFiles) {
            context "$moduleName-`$(`$file.Name)" {
                `$verboseOutputFile = New-TemporaryFile
                `$analysis = Invoke-ScriptAnalyzer -Path "`$(`$file.FullName)" -Recurse -Settings "$SettingsFile" -Verbose -ErrorAction 0 4>"`$verboseOutputFile"
                `$verboseOutput = Get-Content "`$(`$verboseOutputFile.FullName)"
                `$scriptAnalyzerRules = [regex]::matches(`$verboseOutput, '(?<=Running\s)(\w+?)(?=\s)').value | Select-Object -Unique
                foreach (`$rule in `$scriptAnalyzerRules) {
                    it "`$rule" {
                        if (`$analysis.RuleName -contains `$rule) {
                            `$analysis | Where RuleName -eq `$rule -outvariable failures | Out-Default
                            `$failures.Count | Should Be 0
                        }
                    }
                }
            }
        }
    }
"@
    Set-Content -Path $testFilePS1 -Value $testFileContent
    $pssaResults = Invoke-Pester -Path $testFilePS1 -OutputFile $OutputXML -OutputFormat 'NUnitXml' -Passthru
}

end {
    if ($pssaResults.FailedCount -gt 0) {
        throw "$($pssaResults.FailedCount) tests had at least one failure"
    }
}
