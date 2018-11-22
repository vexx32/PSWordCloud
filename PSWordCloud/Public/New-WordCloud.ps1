using namespace System.Collections.Generic
using namespace System.Drawing
using namespace System.Management.Automation
using namespace System.Numerics

class SizeTransformAttribute : ArgumentTransformationAttribute, Attribute {
    static [Dictionary[string, Size]] $StandardSizes = @{
        '720p'  = [Size]::new(1024, 720)
        '1080p' = [Size]::new(1920, 1080)
        '4K'    = [Size]::new(4096, 2160)
    }

    [object] Transform([EngineIntrinsics]$engineIntrinsics, [object] $inputData) {
        $Size = switch ($inputData) {
            { $_ -is [Size] } {
                $_
                break
            }
            { $_ -is [SizeF] } {
                $_.ToSize()
                break
            }
            { $_ -is [int] -or $_ -is [double] } {
                [Size]::new($_, $_)
                break
            }
            { $_ -in [SizeTransformAttribute]::StandardSizes.Keys } {
                [SizeTransformAttribute]::StandardSizes[$_]
                break
            }
            { $_ -is [string] } {
                if ($_ -match '^(?<Width>\d+)x(?<Height>\d+)(px)?$') {
                    [Size]::new($Matches['Width'], $Matches['Height'])
                    break
                }

                if ($_ -match '^(?<Size>\d+)(px)?$') {
                    [Size]::new($Matches['Size'], $Matches['Size'])
                    break
                }
            }
            default {
                throw [ArgumentTransformationMetadataException]::new("Unable to convert entered value $inputData to a valid Size.")
            }
        }

        $Area = $Size.Height * $Size.Width
        if ($Area -ge 100 * 100 -and $Area -le 20000 * 20000) {
            return $Size
        }
        else {
            throw [ArgumentTransformationMetadataException]::new(
                "Specified size $inputData is either too small to use for an image size, or would exceed GDI+ limitations."
            )
        }
    }
}

function New-WordCloud {
    <#
    .SYNOPSIS
    Creates a word cloud from the input text.

    .DESCRIPTION
    Measures the frequency of use of each word, taking into account plural and similar forms, and creates an image
    with each word's visual size corresponding to the frequency of occurrence in the input text.

    .PARAMETER InputString
    The string data to examine and create the word cloud image from.

    .PARAMETER Path
    The output path of the word cloud.

    .PARAMETER ColorSet
    Define a set of colors to use when rendering the word cloud. Any set of [System.Drawing.KnownColor] values will be
    accepted.

    .PARAMETER MaxColors
    Limit the maximum number of colors from either a standard or custom set that will be used. A random selection of
    this many colors will be used to render the word cloud.

    .PARAMETER FontFamily
    Specify the font family as a string or [FontFamily] value. System.Drawing supports primarily TrueType fonts.

    .PARAMETER FontStyle
    Specify the font style to use for the word cloud.

    .PARAMETER ImageSize
    Specify the image size to use in pixels. This value will be used for both the width and height of the rendered
    image. The image dimensions can be any value between 500 and 20,000px. 4096 will be used by default.

    .PARAMETER DistanceStep
    The number of pixels to increment per radial sweep. Higher values will make the operation quicker, but may reduce
    the effectiveness of the packing algorithm. Lower values will take longer, but will generally ensure a more
    tightly-packed word cloud.

    .PARAMETER RadialGranularity
    The number of radial points at each distance step to check during a single sweep. This value is scaled as the radius
    expands to retain some consistency in the overall step distance as the distance from the center increases.

    .PARAMETER BackgroundColor
    Set the background color of the image. Colors with similar names to the background color are automatically excluded
    from being selected. Specify $null to render a transparent image with the word cloud superimposed.

    .PARAMETER Monochrome
    Use only shades of grey to create the word cloud.

    .PARAMETER OutputFormat
    Specify the output image file format to use.

    .PARAMETER MaxWords
    Specify the maximum number of words to include in the word cloud. 100 is default. If there are fewer unique words
    than the maximum set, all unique words will be rendered.

    .EXAMPLE
    Get-Content .\Words.txt | New-WordCloud -Path .\WordCloud.png

    Generates a word cloud from the words in the specified file, and saves it to the specified image file.

    .NOTES
    Only the top 100 most frequent words will be included in the word cloud by default; typically, words that fall under
    this ranking end up being impossible to render cleanly.
    #>
    [CmdletBinding(DefaultParameterSetName = 'SelectColors')]
    [Alias('wordcloud', 'wcloud')]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline)]
        [Alias('InputString', 'InputObject', 'String', 'Words', 'Document', 'Page')]
        [AllowEmptyString()]
        [string[]]
        $Text,

        [Parameter(Mandatory, Position = 1)]
        [Alias('OutFile', 'ExportPath', 'ImagePath')]
        [ValidateScript(
            { Test-Path -IsValid $_ -PathType Leaf }
        )]
        [string[]]
        $Path,

        [Parameter(ParameterSetName = 'SelectColors')]
        [Alias('ColourSet')]
        [KnownColor[]]
        $ColorSet = [KnownColor].GetEnumNames(),

        [Parameter()]
        [Alias('MaxColours')]
        [int]
        $MaxColors,

        [Parameter()]
        [Alias('FontFace')]
        [ArgumentCompleter(
            {
                [FontFamily]::Families.Name.Where{-not [string]::IsNullOrWhiteSpace($_)}.ForEach{
                    if ($_ -match '[\s ]') { "'$_'" }
                    else { $_ }
                }
            }
        )]
        [FontFamily]
        $FontFamily = [FontFamily]::new('Consolas'),

        [Parameter()]
        [FontStyle]
        $FontStyle = [FontStyle]::Regular,

        [Parameter()]
        [Alias('ImagePixelSize')]
        [ArgumentCompleter(
            {
                '720p', '1080p', '4K', '640x1146', '480x800'
            }
        )]
        [SizeTransform()]
        [Size]
        $ImageSize = [Size]::new(4096, 2160),

        [Parameter()]
        [ValidateRange(1, 500)]
        $DistanceStep = 5,

        [Parameter()]
        [ValidateRange(1, 50)]
        $RadialGranularity = 15,

        [Parameter()]
        [Alias('BackgroundColour')]
        [Nullable[KnownColor]]
        $BackgroundColor = [KnownColor]::Black,

        [Parameter(Mandatory, ParameterSetName = 'Monochrome')]
        [Alias('Greyscale', 'Grayscale')]
        [switch]
        $Monochrome,

        [Parameter()]
        [Alias('ImageFormat', 'Format')]
        [ValidateSet("Bmp", "Emf", "Exif", "Gif", "Jpeg", "Png", "Tiff", "Wmf")]
        [string]
        $OutputFormat = "Png",

        [Parameter()]
        [Alias('MaxWords')]
        [ValidateRange(10, 500)]
        [int]
        $MaxUniqueWords = 100
    )
    begin {
        $ExcludedWords = @(
            'a', 'about', 'above', 'after', 'again', 'against', 'all', 'also', 'am', 'an', 'and', 'any', 'are', 'aren''t', 'as',
            'at', 'be', 'because', 'been', 'before', 'being', 'below', 'between', 'both', 'but', 'by', 'can', 'can''t',
            'cannot', 'com', 'could', 'couldn''t', 'did', 'didn''t', 'do', 'does', 'doesn''t', 'doing', 'don''t', 'down',
            'during', 'each', 'else', 'ever', 'few', 'for', 'from', 'further', 'get', 'had', 'hadn''t', 'has', 'hasn''t',
            'have', 'haven''t', 'having', 'he', 'he''d', 'he''ll', 'he''s', 'hence', 'her', 'here', 'here''s', 'hers',
            'herself', 'him', 'himself', 'his', 'how', 'how''s', 'however', 'http', 'i', 'i''d', 'i''ll', 'i''m', 'i''ve', 'if',
            'in', 'into', 'is', 'isn''t', 'it', 'it''s', 'its', 'itself', 'just', 'k', 'let''s', 'like', 'me', 'more', 'most',
            'mustn''t', 'my', 'myself', 'no', 'nor', 'not', 'of', 'off', 'on', 'once', 'only', 'or', 'other', 'otherwise',
            'ought', 'our', 'ours', 'ourselves', 'out', 'over', 'own', 'r', 'same', 'shall', 'shan''t', 'she', 'she''d',
            'she''ll', 'she''s', 'should', 'shouldn''t', 'since', 'so', 'some', 'such', 'than', 'that', 'that''s', 'the',
            'their', 'theirs', 'them', 'themselves', 'then', 'there', 'there''s', 'therefore', 'these', 'they', 'they''d',
            'they''ll', 'they''re', 'they''ve', 'this', 'those', 'through', 'to', 'too', 'under', 'until', 'up', 'very', 'was',
            'wasn''t', 'we', 'we''d', 'we''ll', 'we''re', 'we''ve', 'were', 'weren''t', 'what', 'what''s', 'when', 'when''s',
            'where', 'where''s', 'which', 'while', 'who', 'who''s', 'whom', 'why', 'why''s', 'with', 'won''t', 'would',
            'wouldn''t', 'www', 'you', 'you''d', 'you''ll', 'you''re', 'you''ve', 'your', 'yours', 'yourself', 'yourselves'
        ) -join '|'
        $SplitChars = " `n.,`"?!{}[]:()`“`”™" -as [char[]]
        $ColorIndex = 0
        $RadialDistance = 0

        $WordList = [List[string]]::new()
        $WordHeightTable = @{}
        $WordSizeTable = @{}

        $ExportFormat = $OutputFormat | Get-ImageFormat

        if ($PSCmdlet.ParameterSetName -eq 'Monochrome') {
            $MinSaturation = 0
        }
        else {
            $MinSaturation = 0.5
        }

        if (-not $PSBoundParameters.ContainsKey('MaxColors')) {
            $MaxColors = [int]::MaxValue
        }

        $PathList = foreach ($FilePath in $Path) {
            if ($FilePath -notmatch "\.$OutputFormat$") {
                $FilePath += $OutputFormat
            }
            if (-not (Test-Path -Path $FilePath)) {
                (New-Item -ItemType File -Path $FilePath).FullName
            }
            else {
                (Get-Item -Path $FilePath).FullName
            }
        }

        Write-Verbose "Export Format: $ExportFormat"

        $ColorList = $ColorSet |
            Sort-Object {Get-Random} |
            Select-Object -First $MaxColors |
            ForEach-Object {
            if (-not $Monochrome) {
                [Color]::FromKnownColor($_)
            }
            else {
                [int]$Brightness = [Color]::FromKnownColor($_).GetBrightness() * 255
                [Color]::FromArgb($Brightness, $Brightness, $Brightness)
            }
        } |
            Where-Object {
            if ($BackgroundColor) {
                $_.Name -notmatch $BackgroundColor -and
                $_.GetSaturation() -ge $MinSaturation
            }
            else {
                $_.GetSaturation() -ge $MinSaturation
            }
        } |
            Sort-Object -Descending {
            $Value = $_.GetBrightness()
            $Random = (-$Value..$Value | Get-Random) / (1 - $_.GetSaturation())
            $Value + $Random
        }
    }
    process {
        $WordList.AddRange(
            $Text.Split($SplitChars, [StringSplitOptions]::RemoveEmptyEntries).Where{
                $_ -notmatch "^($ExcludedWords)s?$|^[^a-z]+$|[^a-z0-9'_-]" -and $_.Length -gt 1
            } -replace "^'|'$" -as [string[]]
        )
    }
    end {
        # Count occurrence of each word
        $RemoveList = [List[string]]::new()
        switch ($WordList) {
            { $WordHeightTable[($_ -replace 's$')] } {
                $WordHeightTable[($_ -replace 's$')] += 1
                continue
            }
            { $WordHeightTable[("${_}s")] } {
                $WordHeightTable[($_)] = $WordHeightTable[("${_}s")] + 1
                $RemoveList.Add("${_}s")
                continue
            }
            default {
                $WordHeightTable[($_)] += 1
                continue
            }
        }
        foreach ($Word in $RemoveList) {
            $WordHeightTable.Remove($Word)
        }

        $SortedWordList = $WordHeightTable.GetEnumerator().Name |
            Sort-Object -Descending { $WordHeightTable[$_] } |
            Select-Object -First $MaxUniqueWords

        $HighestFrequency, $AverageFrequency = $SortedWordList |
            ForEach-Object { $WordHeightTable[$_] } |
            Measure-Object -Average -Maximum |
            ForEach-Object {$_.Maximum, $_.Average}

        $FontScale = ($ImageSize.Height + $ImageSize.Width) / ($AverageFrequency * $SortedWordList.Count)

        Write-Verbose "Unique Words Count: $($WordHeightTable.PSObject.BaseObject.Count)"
        Write-Verbose "Highest Word Frequency: $HighestFrequency; Average: $AverageFrequency"
        Write-Verbose "Max Font Size: $($HighestFrequency * $FontScale)"

        try {
            # Create a graphics object to measure the text's width and height.
            $DummyImage = [Bitmap]::new(1, 1)
            $Graphics = [Graphics]::FromImage($DummyImage)

            foreach ($Word in $SortedWordList) {
                $WordHeightTable[$Word] = [Math]::Round($WordHeightTable[$Word] * $FontScale)
                if ($WordHeightTable[$Word] -lt 8) { continue }

                $Font = [Font]::new(
                    $FontFamily,
                    $WordHeightTable[$Word],
                    $FontStyle,
                    [GraphicsUnit]::Pixel
                )

                $WordSizeTable[$Word] = $Graphics.MeasureString($Word, $Font)
            }

            $WordHeightTable | Out-String | Write-Debug
        }
        catch {
            $PSCmdlet.ThrowTerminatingError($_)
        }
        finally {
            if ($Graphics) {
                $Graphics.Dispose()
            }
            if ($DummyImage) {
                $DummyImage.Dispose()
            }
        }

        $MaxSideLength = [Math]::Max($ImageSize.Width, $ImageSize.Height)
        $AspectRatio = $ImageSize.Width / $ImageSize.Height
        $CentrePoint = [PointF]::new($ImageSize.Width / 2, $ImageSize.Height / 2)
        Write-Verbose "Image dimensions: $ImageSize with centrepoint $CentrePoint and ratio $AspectRatio."

        try {
            $WordCloudImage = [Bitmap]::new($ImageSize.Width, $ImageSize.Height)
            $DrawingSurface = [Graphics]::FromImage($WordCloudImage)

            if ($BackgroundColor) {
                $DrawingSurface.Clear([Color]::FromKnownColor($BackgroundColor))
            }
            $DrawingSurface.SmoothingMode = [Drawing2D.SmoothingMode]::AntiAlias
            $DrawingSurface.TextRenderingHint = [Text.TextRenderingHint]::AntiAlias

            $RectangleList = [List[RectangleF]]::new()
            $RadialScanCount = 0
            $Jitter = [Random]::new()
            :words foreach ($Word in $SortedWordList) {
                if (-not $WordSizeTable[$Word]) { continue }

                $Font = [Font]::new(
                    $FontFamily,
                    $WordHeightTable[$Word],
                    $FontStyle,
                    [GraphicsUnit]::Pixel
                )

                $RadialScanCount /= 3
                $WordRectangle = $null
                do {
                    if ( $RadialDistance -gt ($MaxSideLength / 2) ) {
                        $RadialDistance = $MaxSideLength / $DistanceStep / 25
                        continue words
                    }

                    $AngleIncrement = 360 / ( ($RadialDistance + 1) * $RadialGranularity / 10 )
                    switch ([int]$RadialScanCount -band 7) {
                        0 { $Start = 0; $End = 360 }
                        1 { $Start = -90; $End = 270 }
                        2 { $Start = -180; $End = 180 }
                        3 { $Start = -270; $End = 90  }
                        4 { $Start = 360; $End = 0; $AngleIncrement *= -1 }
                        5 { $Start = 270; $End = -90; $AngleIncrement *= -1 }
                        6 { $Start = 180; $End = -180; $AngleIncrement *= -1 }
                        7 { $Start = 90; $End = -270; $AngleIncrement *= -1 }
                    }

                    $IsColliding = $false
                    for (
                        $Angle = $Start;
                        $( if ($Start -lt $End) {$Angle -le $End} else {$End -le $Angle} );
                        $Angle += $AngleIncrement
                    ) {
                        $Radians = Convert-ToRadians -Degrees $Angle
                        $Complex = [Complex]::FromPolarCoordinates($RadialDistance, $Radians)

                        $OffsetX = $WordSizeTable[$Word].Width * $Jitter.NextDouble()
                        $OffsetY = $WordSizeTable[$Word].Height * $Jitter.NextDouble()
                        $DrawLocation = [PointF]::new(
                            $Complex.Real * $AspectRatio + $CentrePoint.X - $OffsetX,
                            $Complex.Imaginary + $CentrePoint.Y - $OffsetY
                        )

                        $WordRectangle = [RectangleF]::new([PointF]$DrawLocation, [SizeF]$WordSizeTable[$Word])

                        foreach ($Rectangle in $RectangleList) {
                            $IsColliding = (
                                $WordRectangle.IntersectsWith($Rectangle) -or
                                $WordRectangle.Top -lt 0 -or
                                $WordRectangle.Bottom -gt $ImageSize.Height -or
                                $WordRectangle.Left -lt 0 -or
                                $WordRectangle.Right -gt $ImageSize.Width
                            )

                            if ($IsColliding) {
                                break
                            }
                        }

                        if (!$IsColliding) {
                            break
                        }
                    }

                    if ($IsColliding) {
                        $RadialDistance += $DistanceStep
                        $RadialScanCount++
                    }
                } while ($IsColliding)

                $RectangleList.Add($WordRectangle)
                $Color = $ColorList[$ColorIndex]

                $ColorIndex++
                if ($ColorIndex -ge $ColorList.Count) {
                    $ColorIndex = 0
                }

                Write-Debug "Writing $Word with font $Font in $Color at $DrawLocation"
                $DrawingSurface.DrawString($Word, $Font, [SolidBrush]::new($Color), $DrawLocation)

                $RadialDistance -= $DistanceStep * ($RadialScanCount / 2)
            }

            $DrawingSurface.Flush()
            foreach ($FilePath in $PathList) {
                $WordCloudImage.Save($FilePath, $ExportFormat)
            }
        }
        catch {
            $PSCmdlet.ThrowTerminatingError($_)
        }
        finally {
            $DrawingSurface.Dispose()
            $WordCloudImage.Dispose()
        }
    }
}