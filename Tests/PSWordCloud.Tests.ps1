Describe 'PSWordCloud Tests' {

    BeforeAll {
        $File = New-TemporaryFile |
            Rename-Item -NewName { "$($_.BaseName).svg" } -PassThru
    }

    It 'should be able to import the PSWordCloud module successfully' {
        { Import-Module PSWordCloud -ErrorAction Stop } | Should -Not -Throw
    }

    It 'should run New-WordCloud without errors' {
        Get-ChildItem -Path "$PSScriptRoot/../" -Recurse -File -Include "*.cs", "*.ps*1", "*.md" |
            Get-Content |
            New-WordCloud -Path $File.FullName
    }

    It 'should create a new SVG file' {
        $File | Should -Exist
    }

    It 'should create non-empty files' {
        $File = Get-Item -Path $File.FullName
        $File.Length | Should -BeGreaterThan 0
    }
}
