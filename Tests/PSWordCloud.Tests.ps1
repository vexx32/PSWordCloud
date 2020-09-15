Describe 'PSWordCloud Tests' {

    BeforeAll {
        $FileName = New-TemporaryFile |
            Rename-Item -NewName { "$($_.BaseName).svg" } -PassThru |
            Select-Object -ExpandProperty FullName

        Remove-Item -Path $FileName -Force
    }

    It 'should be able to import the PSWordCloud module successfully' {
        { Import-Module PSWordCloud -ErrorAction Stop } | Should -Not -Throw
    }

    It 'should error out for empty or otherwise unusable input' {
        { [string]::Empty | New-WordCloud -Path ./test.svg } | Should -Throw -ExpectedMessage "No usable input was provided. Please provide string data via the pipeline or in a word size dictionary."
    }

    Context 'FileSystem Provider' {
        It 'should run New-WordCloud without errors' {
            Get-ChildItem -Path "$PSScriptRoot/../" -Recurse -File -Include "*.cs", "*.ps*1", "*.md" |
                Get-Content |
                New-WordCloud -Path $FileName
        }

        It 'should create a new SVG file' {
            $FileName | Should -Exist
        }

        It 'should create a non-empty file' {
            $FileName = Get-Item -Path $FileName
            (Get-Item -Path $FileName).Length | Should -BeGreaterThan 0
        }

        It 'should have SVG data in the file' {
            Select-String -Pattern '<svg.*>' -Path $FileName | Should -Not -BeNullOrEmpty
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
