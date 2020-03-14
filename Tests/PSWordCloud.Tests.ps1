Describe 'PSWordCloud Tests' {

    BeforeAll {
        $File = New-TemporaryFile |
            Rename-Item -NewName { "$($_.BaseName).svg" } -PassThru
    }

    It 'should be able to import the PSWordCloud module successfully' {
        { Import-Module PSWordCloud -ErrorAction Stop } | Should -Not -Throw
    }

    Context 'FileSystem Provider' {
        It 'should run New-WordCloud without errors' {
            Get-ChildItem -Path "$PSScriptRoot/../" -Recurse -File -Include "*.cs", "*.ps*1", "*.md" |
                Get-Content |
                New-WordCloud -Path $File.FullName
        }

        It 'should create a new SVG file' {
            $File | Should -Exist
        }

        It 'should create a non-empty file' {
            $File = Get-Item -Path $File.FullName
            $File.Length | Should -BeGreaterThan 0
        }

        It 'should have SVG data in the file' {
            Select-String -Pattern '<svg.*>'  -Path $File.FullName | Should -Not -BeNullOrEmpty
        }
    }

    Context 'Variable Provider' {

        BeforeAll {
            $VariableName = 'SvgData'
        }

        It 'should run New-WordCloud without errors' {
            Get-ChildItem -Path "$PSScriptRoot/../" -Recurse -File -Include "*.cs", "*.ps*1", "*.md" |
                Get-Content |
                New-WordCloud -Path "variable:global:$VariableName"
        }

        It 'should create a new variable' {
            { Get-Variable -Name $VariableName -ErrorAction Stop } | Should -Not -Throw
        }

        It 'should populate the variable' {
            (Get-Variable -Name $VariableName).Value.Length | Should -BeGreaterThan 0
        }

        It 'should populate the variable with SVG data' {
            (Get-Variable -Name $VariableName).Value | Should -MatchExactly '<svg.*>'
        }
    }
}
