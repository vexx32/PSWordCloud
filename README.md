# PSWordCloud

Create pretty word clouds with PowerShell!

## Installation

```powershell
Install-Module PSWordCloud
```

### Linux & macOS

To get it working on Linux or macOS, you must install the `mono-libgdiplus` package on your machine for this to work.

For macOS, you can use Homebrew:

```sh
brew install mono-libgdiplus
```

For linux, use your package manager of choice to install it. For example, via Apt:

```sh
apt install libgdiplus
```

## Usage

```powershell
Get-ClipBoard | New-WordCloud -Path .\wordcloud.png -FontFamily Georgia

Get-Content .\words.txt | New-WordCloud -Path .\wordcloud2.png -ImageSize 1080p

Get-ClipBoard | New-WordCloud -Path .\wordcloud3.png -FontFamily Consolas -ImageSize 1000x1000
```

## Examples

See [the Gallery](./Examples/Examples.md) for some example usage and output!