# PSWordCloud

Create pretty word clouds with PowerShell!

## Installation

```powershell
Install-Module PSWordCloud
```

## Usage

```powershell

# Simply provide a list of words (in this case, supplied with a simple hashtable depicting words and
# their relative sizes.

New-WordCloud -Path .\wordcloud.svg -Typeface Consolas -WordSizes @{
    dragon = Get-Random -Maximum 10 -Minimum 1 
    rabbit = Get-Random -Maximum 15 -Minimum 1 
    horse = Get-Random -Maximum 18 -Minimum 1 
    cow = Get-Random -Maximum 20 -Minimum 1 
    cat = Get-Random -Maximum 8 -Minimum 1 
    fox = Get-Random -Maximum 12 -Minimum 1 
}

# Alternately, get a chunk of text (doesn't matter where), and pipe it directly to the cmdlet to create
# a word-frequency word cloud.
Get-ClipBoard | New-WordCloud -Path .\wordcloud.svg -Typeface Georgia

Get-Content .\words.txt | New-WordCloud -Path .\wordcloud2.svg -ImageSize 1080p

Get-ClipBoard | New-WordCloud -Path .\wordcloud3.svg -Typeface Consolas -ImageSize 1000x1000
```

## Examples

See [the Gallery](./Examples/Examples.md) for some example usage and output!
