# PSWordCloud

Create pretty word clouds with PowerShell!

## Installation

```powershell
Install-Module PSWordCloud
```

## Usage

```powershell
Get-ClipBoard | New-WordCloud -Path .\wordcloud.svg -Typeface Georgia

Get-Content .\words.txt | New-WordCloud -Path .\wordcloud2.svg -ImageSize 1080p

Get-ClipBoard | New-WordCloud -Path .\wordcloud3.svg -Typeface Consolas -ImageSize 1000x1000
```

## Examples

See [the Gallery](./Examples/Examples.md) for some example usage and output!
