# PSWordCloud

Create pretty word clouds with PowerShell!

## Installation

```powershell
Install-Module PSWordCloud
```

## Usage

```powershell

#Simply Provide a list of words (in this case, randomly adding picking one hundred animals)
$animals = New-Object System.Collections.ArrayList
1..100 | %{
    $x = get-random -Maximum 7 -Minimum 1 
    $y = switch ($x){
        1 {"dragon"}
        2 {"rabbit"}
        3{"horse"}
        4{"cow"}
        5{"cat"}
        6{"fox"}
        }

    [void]$animals.add( $y )
}

$animals | New-WordCloud -Path .\wordcloud.svg -Typeface Consolas

Get-ClipBoard | New-WordCloud -Path .\wordcloud.svg -Typeface Georgia

Get-Content .\words.txt | New-WordCloud -Path .\wordcloud2.svg -ImageSize 1080p

Get-ClipBoard | New-WordCloud -Path .\wordcloud3.svg -Typeface Consolas -ImageSize 1000x1000
```

## Examples

See [the Gallery](./Examples/Examples.md) for some example usage and output!
