$Lines = '-' * 70
$PSVersion = $PSVersionTable.PSVersion

Write-Host $Lines
Write-Host "STATUS: Testing with PowerShell v$($PSVersionTable.PSVersion)"
Write-Host $Lines

Import-Module 'PSWordCloud'

$Timestamp = Get-Date -Format "yyyyMMdd-hhmmss"
$TestFile = "PS${PSVersion}_${TimeStamp}_PSWordCloud.TestResults.xml"


# Tell Azure where the test results & code coverage files will be
Write-Host "##vso[task.setvariable variable=TestResults]$TestFile"
Write-Host "##vso[task.setvariable variable=SourceFolders]$ModuleFolders"

# Gather test results. Store them in a variable and file
$PesterParams = @{
    Path         = "$env:PROJECTROOT/Tests"
    PassThru     = $true
    OutputFormat = 'NUnitXml'
    OutputFile   = "$env:BUILD_ARTIFACTSTAGINGDIRECTORY/$TestFile"
    Show         = "Header", "Failed", "Summary"
}
$TestResults = Invoke-Pester @PesterParams

# If tests failed, write errors and exit
if ($TestResults.FailedCount -gt 0) {
    Write-Error "Failed $($TestResults.FailedCount) tests; build failed!"
    exit $TestResults.FailedCount
}
