Describe 'PSWordCloud Tests' {

    BeforeAll {
        $FilePath = Join-Path $env:TEMP -ChildPath "$(New-Guid).svg"
    }

    It 'should be able to import the PSWordCloud module successfully' {
        { Import-Module PSWordCloud -ErrorAction Stop } | Should -Not -Throw
    }

    It 'should run New-WordCloud without errors' {
        Get-ChildItem -Path "$PSScriptRoot/../" -Recurse -File -Include "*.cs", "*.ps*1", "*.md" |
            Get-Content |
            New-WordCloud -Path $FilePath.FullName
    }

    It 'should create a new SVG file' {
        $FilePath | Should -Exist
    }

    It 'should create non-empty files' {
        $File = Get-Item -Path $FilePath.FullName
        $File.Length | Should -BeGreaterThan 0
    }
}
