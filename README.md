# PSWordCloud

Create pretty word clouds with PowerShell!

## Installation

```powershell
Install-Module PSWordCloud
```

## Usage

```powershell
Get-ClipBoard | New-WordCloud -Path .\wordcloud.png -FontFamily Georgia

Get-Content .\words.txt | New-WordCloud -Path .\wordcloud2.png -ImageSize 1080p

Get-ClipBoard | New-WordCloud -Path .\wordcloud3.png -FontFamily Consolas -ImageSize 1000x1000
```