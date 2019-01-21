using namespace System.Collections.Generic
using namespace System.Drawing
using namespace System.Drawing.Drawing2D
using namespace System.IO
using namespace System.Management.Automation
using namespace System.Management.Automation.Runspaces
using namespace System.Numerics

class SizeTransformAttribute : ArgumentTransformationAttribute {
    static [hashtable] $StandardSizes = @{
        '720p'  = [Size]::new(1280, 720)
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
                if ($_ -match '^(?<Width>[\d\.,]+)x(?<Height>[\d\.,]+)(px)?$') {
                    [Size]::new($Matches['Width'], $Matches['Height'])
                    break
                }

                if ($_ -match '^(?<Size>[\d\.,]+)(px)?$') {
                    [Size]::new($Matches['Size'], $Matches['Size'])
                    break
                }
            }
            default {
                throw [ArgumentTransformationMetadataException]::new("Unable to convert entered value $inputData to a valid [System.Drawing.Size].")
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

class ColorTransformAttribute : ArgumentTransformationAttribute {
    static [string[]] $ColorNames = @(
        [KnownColor].GetEnumNames()
        "Transparent"
    )

    [object] Transform([EngineIntrinsics]$engineIntrinsics, [object] $inputData) {
        $Items = switch ($inputData) {
            { $_ -eq $null -or $_ -eq 'Transparent' } {
                [Color]::Transparent
                continue
            }
            { $_ -as [KnownColor] } {
                [Color]::FromKnownColor($_ -as [KnownColor])
                continue
            }
            { $_ -is [Color] } {
                $_
                continue
            }
            { $_ -is [string] } {
                if ($_ -match 'R(?<Red>[0-9]{1,3})G(?<Green>[0-9]{1,3})B(?<Blue>[0-9]{1,3})') {
                    [Color]::FromArgb($Matches['Red'], $Matches['Green'], $Matches['Blue'])
                    continue
                }

                if ($_ -match 'R(?<Red>[0-9]{1,3})G(?<Green>[0-9]{1,3})B(?<Blue>[0-9]{1,3})A(?<Alpha>[0-9]{1,3})') {
                    [Color]::FromArgb($Matches['Alpha'], $Matches['Red'], $Matches['Green'], $Matches['Blue'])
                    continue
                }

                if ($MatchingValues = [KnownColor].GetEnumNames() -like $_) {
                    ($MatchingValues -as [KnownColor[]]).ForEach{ [Color]::FromKnownColor($_) }
                }
            }
            { $_ -is [int] } {
                [Color]::FromArgb($_)
                continue
            }
            default {
                throw [ArgumentTransformationMetadataException]::new("Could not convert value '$_' to a valid [System.Drawing.Color] or [System.Drawing.KnownColor].")
            }
        }

        return $Items
    }
}

class FileTransformAttribute : ArgumentTransformationAttribute {
    [object] Transform([EngineIntrinsics]$engineIntrinsics, [object] $inputData) {
        $Items = switch ($inputData) {
            { $_ -as [FileInfo] } {
                $_
                break
            }
            { $_ -is [string] } {
                $Path = Resolve-Path -Path $_
                if (@($Path).Count -gt 1) {
                    throw [ArgumentTransformationMetadataException]::new("Multiple files found, please enter only one: $($Path -join ', ')")
                }

                if (Test-Path -Path $Path -PathType Leaf) {
                    [FileInfo]::new($Path)
                }

                break
            }
            default {
                throw [ArgumentTransformationMetadataException]::new("Could not convert value '$_' to a valid [System.IO.FileInfo] object.")
            }
        }

        return $Items
    }
}

function New-WordCloud {
    <#
    .SYNOPSIS
    Creates a word cloud from the input text.

    .DESCRIPTION
    Measures the frequency of use of each word, taking into account plural and similar forms, and creates an image
    with each word's visual size corresponding to the frequency of occurrence in the input text.

    .PARAMETER InputObject
    The string data to examine and create the word cloud image from. Any non-string data piped in will be passed through
    Out-String first to obtain a proper string representation of the object, which will then be broken down into its
    constituent words.

    .PARAMETER Path
    The output path of the word cloud.

    .PARAMETER ColorSet
    Define a set of colors to use when rendering the word cloud. Any array of values in any mix of the following formats
    is acceptable:

    - Strings denoting a color name, or a wildcarded color name that matches one or more colors, e.g., *blue*
    - Valid [System.Drawing.Color] objects
    - Valid [System.Drawing.KnownColor] values in enum or string format
    - Strings of the format r255g255b255 or r255g255b255a255 where the integers are the R, G, B, and optionally Alpha
      values of the desired color.
    - Any valid integer value; these are passed directly to [System.Drawing.Color]::FromArgb($Integer) to be converted
      into valid colors.

    .PARAMETER MaxColors
    Limit the maximum number of colors from either the standard or custom set that will be used. A random selection of
    this many colors will be used to render the word cloud.

    .PARAMETER FocusWord
    Specifies a title or centred focus word to insert into the word cloud. This word will be rendered in the middle of
    the cloud, slightly larger than the largest word in the cloud.

    .PARAMETER ExcludeWord
    Specifies one or more words to exclude from the word cloud.

    .PARAMETER FontFamily
    Specify the font family as a string or [FontFamily] value.

    .PARAMETER FontStyle
    Specify the font style to use for the word cloud.

    .PARAMETER StrokeColor
    Specify the color of outline to use when rendering words.

    .PARAMETER StrokeWidth
    Specify the width of the outline to use use when rendering words. Off by default.

    .PARAMETER ImageSize
    Specify the image size to use in pixels. The image dimensions can be any value between 500 and 20,000px. Any of the
    following size specifier formats are permitted:

    - Any valid [System.Drawing.Size] object
    - Any valid [System.Drawing.SizeF] object
    - 1000x1000
    - 1000x1000px
    - 1000
    - 1000px
    - 720p	        (Creates an image of size 1280x720px)
    - 1080p         (Creates an image of size 1920x1080px)
    - 4K	        (Creates an image of size 4096x2160px)

    4096x2160 will be used by default. Note that the minimum image resolution is 10,000 pixels (100 x 100px), and the
    maximum resolution is 400,000,000 pixels (20,000 x 20,000px, 400MP).

    .PARAMETER DistanceStep
    The number of pixels to increment per radial sweep. Higher values will make the operation quicker, but may reduce
    the effectiveness of the packing algorithm. Lower values will take longer, but will generally ensure a more
    tightly-packed word cloud.

    .PARAMETER RadialGranularity
    The number of radial points at each distance step to check during a single sweep. This value is scaled as the radius
    expands to retain some consistency in the overall step distance as the distance from the center increases.

    .PARAMETER BackgroundColor
    Set the background color of the image. Colors with similar names to the background color are automatically excluded
    from being selected for use in word coloring. Any value in of the following formats is acceptable:

    - Valid [System.Drawing.Color] objects
    - Valid [System.Drawing.KnownColor] values in enum or string format
    - Strings of the format r255g255b255 or r255g255b255a255 where the integers are the R, G, B, and optionally Alpha
      values of the desired color.
    - Any valid integer value; these are passed directly to [System.Drawing.Color]::FromArgb($Integer) to be converted
      into valid colors.

    Specify $null or Transparent as the background color value to render the word cloud on a transparent background.

    .PARAMETER Monochrome
    Use only shades of grey to create the word cloud.

    .PARAMETER OutputFormat
    Specify the output image file format to use.

    .PARAMETER BackgroundImage
    Specify a base image to superimpose the word cloud on. The image will not be resized.

    .PARAMETER MaxUniqueWords
    Specify the maximum number of words to include in the word cloud. 100 is default. If there are fewer unique words
    than the maximum amount, all unique words will be rendered.

    .PARAMETER RandomSeed
    Specify an integer seed value to generate randomness from. Identical seed and input values should render comparable
    word clouds.

    .PARAMETER Padding
    Specify the bounded spacing between words. This is calculated as additional distance from the prior drawn paths and
    each new word placement attempt. Higher values will tend to take longer to render, but tend to look neater. 3.5 is
    the default. Specify 0 for no minimum spacing and negative values to allow words to overlap.

    .PARAMETER WordScale
    Adjust the word scaling. 1 is the default.

    .PARAMETER AllowOverflow
    Toggles whether words are permitted to be drawn partially crossing the boundary of the iamge.

    .PARAMETER DisableWordRotation
    Disables drawing words vertically.

    .PARAMETER BoxCollisions
    Reverts collision detection to the slightly faster but more boxy detection methods based purely on path bounds
    rather than the path shape itself.

    .PARAMETER AllowStopWords
    Removes the standard ignored words from the ignore list. The standard ignore list ignores single-letter words and
    some of the most common shorter English words.

    .EXAMPLE
    Get-Content .\Words.txt | New-WordCloud -Path .\WordCloud.png

    Generates a word cloud from the words in the specified file, and saves it to the specified image file.

    .NOTES
    Only the top 100 most frequent words will be included in the word cloud by default; typically, words that fall under
    this ranking end up being impossible to render cleanly except on very high resolutions.

    The word cloud will be rendered according to the image size; landscape or portrait configurations will result in
    ovoid clouds, whereas square images will result mainly in circular clouds.
    #>
    [CmdletBinding(DefaultParameterSetName = 'ColorBackground', SupportsShouldProcess)]
    [Alias('wordcloud', 'wcloud')]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline, ParameterSetName = 'ColorBackground')]
        [Parameter(Mandatory, Position = 0, ValueFromPipeline, ParameterSetName = 'ColorBackground-Mono')]
        [Parameter(Mandatory, Position = 0, ValueFromPipeline, ParameterSetName = 'FileBackground')]
        [Parameter(Mandatory, Position = 0, ValueFromPipeline, ParameterSetName = 'FileBackground-Mono')]
        [Alias('InputString', 'Text', 'String', 'Words', 'Document', 'Page')]
        [AllowEmptyString()]
        [object[]]
        $InputObject,

        [Parameter(Mandatory, Position = 1, ParameterSetName = 'ColorBackground')]
        [Parameter(Mandatory, Position = 1, ParameterSetName = 'ColorBackground-Mono')]
        [Parameter(Mandatory, Position = 1, ParameterSetName = 'FileBackground')]
        [Parameter(Mandatory, Position = 1, ParameterSetName = 'FileBackground-Mono')]
        [Alias('OutFile', 'ExportPath', 'ImagePath')]
        [ValidateScript(
            { Test-Path -IsValid $_ -PathType Leaf }
        )]
        [string[]]
        $Path,

        [Parameter(ParameterSetName = 'ColorBackground')]
        [Parameter(ParameterSetName = 'ColorBackground-Mono')]
        [Parameter(ParameterSetName = 'FileBackground')]
        [Parameter(ParameterSetName = 'FileBackground-Mono')]
        [Alias('ColourSet')]
        [ArgumentCompleter(
            {
                param($Command, $Parameter, $WordToComplete, $CommandAst, $FakeBoundParams)
                if (!$WordToComplete) {
                    return [ColorTransformAttribute]::ColorNames
                }
                else {
                    return [ColorTransformAttribute]::ColorNames.Where{ $_ -match "^$WordToComplete" }
                }
            }
        )]
        [ColorTransformAttribute()]
        [Color[]]
        $ColorSet = [ColorTransformAttribute]::ColorNames,

        [Parameter()]
        [Alias('MaxColours')]
        [int]
        $MaxColors = [int]::MaxValue,

        [Parameter()]
        [Alias('Title')]
        [string]
        $FocusWord,

        [Parameter()]
        [Alias('IgnoreWord')]
        [string[]]
        $ExcludeWord,

        [Parameter()]
        [Alias('FontFace')]
        [ArgumentCompleter(
            {
                param($Command, $Parameter, $WordToComplete, $CommandAst, $FakeBoundParams)
                $FontLibrary = [FontFamily]::Families.Name.Where{-not [string]::IsNullOrWhiteSpace($_)}

                if (!$WordToComplete) {
                    return $FontLibrary -replace '(?="|`|\$)', '`' -replace '^|$', '"'
                }
                else {
                    return $FontLibrary.Where{$_ -match "^('|`")?$([regex]::Escape($WordToComplete))"} -replace '(?="|`|\$)', '`' -replace '^|$', '"'
                }
            }
        )]
        [FontFamily]
        $FontFamily = [FontFamily]::new('Consolas'),

        [Parameter()]
        [FontStyle]
        $FontStyle = [FontStyle]::Regular,

        [Parameter()]
        [Alias('StrokeColour')]
        [ArgumentCompleter(
            {
                param($Command, $Parameter, $WordToComplete, $CommandAst, $FakeBoundParams)
                if (!$WordToComplete) {
                    return [ColorTransformAttribute]::ColorNames
                }
                else {
                    return [ColorTransformAttribute]::ColorNames.Where{ $_ -match "^$WordToComplete" }
                }
            }
        )]
        [ColorTransformAttribute()]
        [Color]
        $StrokeColor = [Color]::Black,

        [Parameter()]
        [Alias('Outline')]
        [double]
        $StrokeWidth,

        [Parameter(ParameterSetName = 'ColorBackground')]
        [Parameter(ParameterSetName = 'ColorBackground-Mono')]
        [Alias('ImagePixelSize')]
        [ArgumentCompleter(
            {
                param($Command, $Parameter, $WordToComplete, $CommandAst, $FakeBoundParams)

                $Values = @('720p', '1080p', '4K', '640x1146', '480x800')

                if ($WordToComplete) {
                    return $Values.Where{$_ -match "^$WordToComplete"}
                }
                else {
                    return $Values
                }
            }
        )]
        [SizeTransformAttribute()]
        [Size]
        $ImageSize = [Size]::new(4096, 2160),

        [Parameter()]
        [ValidateRange(1, 500)]
        $DistanceStep = 5,

        [Parameter()]
        [ValidateRange(1, 50)]
        $RadialGranularity = 15,

        [Parameter(ParameterSetName = 'ColorBackground')]
        [Parameter(ParameterSetName = 'ColorBackground-Mono')]
        [Alias('BackgroundColour')]
        [ArgumentCompleter(
            {
                param($Command, $Parameter, $WordToComplete, $CommandAst, $FakeBoundParams)
                if (!$WordToComplete) {
                    return [ColorTransformAttribute]::ColorNames
                }
                else {
                    return [ColorTransformAttribute]::ColorNames.Where{ $_ -match "^$WordToComplete" }
                }
            }
        )]
        [ColorTransformAttribute()]
        [Color]
        $BackgroundColor = [Color]::Black,

        [Parameter(Mandatory, ParameterSetName = 'ColorBackground-Mono')]
        [Parameter(Mandatory, ParameterSetName = 'FileBackground-Mono')]
        [Alias('Greyscale', 'Grayscale')]
        [switch]
        $Monochrome,

        [Parameter()]
        [Alias('ImageFormat', 'Format')]
        [ValidateSet('Bmp', 'Emf', 'Exif', 'Gif', 'Jpeg', 'Png', 'Tiff', 'Wmf')]
        [string]
        $OutputFormat = 'Png',

        [Parameter(Mandatory, ParameterSetName = 'FileBackground')]
        [Parameter(Mandatory, ParameterSetName = 'FileBackground-Mono')]
        [Alias('BaseImage')]
        [FileTransformAttribute()]
        [FileInfo]
        $BackgroundImage,

        [Parameter()]
        [Alias('MaxWords')]
        [ValidateRange(0, 1000)]
        [int]
        $MaxUniqueWords = 100,

        [Parameter()]
        [Alias('Seed')]
        [int]
        $RandomSeed,

        [Parameter()]
        [Alias('Spacing')]
        [double]
        $Padding = 3.5,

        [Parameter()]
        [Alias('Scale', 'FontScale')]
        [double]
        $WordScale = 1,

        [Parameter()]
        [Alias('DisableClipping', 'NoClip')]
        [switch]
        $AllowOverflow,

        [Parameter()]
        [Alias('DisableRotation', 'NoRotation')]
        [switch]
        $DisableWordRotation,

        [Parameter()]
        [Alias('Boxy')]
        [switch]
        $BoxCollisions,

        [Parameter()]
        [switch]
        $AllowStopWords
    )
    begin {
        Write-Debug "Color set: $($ColorSet -join ', ')"
        Write-Debug "Background color: $BackgroundColor"

        $StopWordsPattern = @(
            switch ($true) {
                (-not $AllowStopWords) {
                    (Get-Content "$script:ModuleRoot/Data/StopWords.txt")
                }
                ([bool] $ExcludeWord) {
                    $ExcludeWord
                }
            }
        ).ForEach{[Regex]::Escape($_)} -join '|'
        $SplitChars = " .,`"?!{}[]:()`“`”™*#%^&+=" -as [char[]]
        $ColorIndex = 0
        $RadialDistance = 0
        $GraphicsUnit = [GraphicsUnit]::Pixel

        $WordList = [List[string]]::new(128)
        $WordHeightTable = @{}
        $WordSizeTable = @{}

        $RNG = if ($PSBoundParameters.ContainsKey('RandomSeed')) {
            [Random]::new($RandomSeed)
        }
        else {
            [Random]::new()
        }

        $ProgressID = $RNG.Next()

        $ExportFormat = $OutputFormat | Get-ImageFormat

        if ($PSCmdlet.ParameterSetName -eq 'Monochrome') {
            $MinSaturation = 0
        }
        else {
            $MinSaturation = 0.5
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

        $ColorList = $ColorSet |
            Sort-Object {Get-Random} |
            Select-Object -First $MaxColors |
            ForEach-Object {
            if (-not $Monochrome) {
                $_
            }
            else {
                [int]$Brightness = $_.GetBrightness() * 255
                [Color]::FromArgb($Brightness, $Brightness, $Brightness)
            }
        } | Where-Object {
            if ($BackgroundColor) {
                $_.Name -notmatch $BackgroundColor -and
                $_.GetSaturation() -ge $MinSaturation
            }
            else {
                $_.GetSaturation() -ge $MinSaturation
            }
        } | Sort-Object -Descending {
            $Value = $_.GetBrightness()
            $Random = (-$Value..$Value | Get-Random) / (1 - $_.GetSaturation())
            $Value + $Random
        }

        $RunspacePool = [RunspaceFactory]::CreateRunspacePool(1, 15)
        [List[PSCustomObject]] $RSJobs = [List[PSCustomObject]]::new()

        $InputProcessingScript = {
            param($Strings, $SplitChars, $ExcludedWords)

            foreach ($Item in $Strings) {
                if ($Item -isnot [string]) {
                    [string] $Item = $Item | Out-String
                }

                @($Item -split "\r?\n").Split($SplitChars, [StringSplitOptions]::RemoveEmptyEntries).Where{
                    $_ -notmatch "^($ExcludedWords)s?$|^[^a-z]+$|[^a-z0-9'_-]" -and $_.Length -gt 1
                } -replace "^('|_)|('|_)$"
            }
        }

        $RunspacePool.Open()
        $JobsStarted = 0
        $JobsReceived = 0
        $LineCount = 0
        $WordCount = 0

        [List[object]] $InputStorage = [List[object]]::new()
    }
    process {
        try {
            $LineCount += @($InputObject).Count

            if ($MyInvocation.ExpectingInput) {
                $InputStorage.AddRange($InputObject)
            }
            else {
                $InputStorage = [List[object]] $InputObject
            }

            if ($InputStorage.Count -ge 2000) {
                $WordCount = $WordList.Count
                $PowerShell = [PowerShell]::Create().AddScript($InputProcessingScript).AddParameters(
                    @{
                        Strings       = $InputStorage.ToArray()
                        SplitChars    = $SplitChars
                        ExcludedWords = $StopWordsPattern
                    }
                )
                $PowerShell.RunspacePool = $RunspacePool

                $ProgressParams = @{
                    Activity         = "Processing Input Items: $LineCount"
                    Status           = "Jobs: <Run> $JobsStarted <Completed> $JobsReceived"
                    Id               = $ProgressID
                    CurrentOperation = "Splitting text into words. Word count: $WordCount"
                }
                Write-Progress @ProgressParams

                $RSJobs.Add(
                    [PSCustomObject]@{
                        Instance = $PowerShell
                        Result   = $PowerShell.BeginInvoke()
                    }
                )
                $JobsStarted ++
                $InputStorage.Clear()
            }

            if ($JobsStarted % 15 -eq 1) {
                foreach ($Runspace in $RSJobs.Where{$_.Result.IsCompleted}) {
                    [string[]] $Result = $Runspace.Instance.EndInvoke($Runspace.Result).ForEach{$_}
                    $WordList.AddRange($Result)
                    $RSJobs.Remove($Runspace) > $null
                    $JobsReceived ++
                }
            }
        }
        catch {
            $RunspacePool.Close()
            $RunspacePool.Dispose()
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
    end {
        try {
            if ($InputStorage.Count -gt 0) {
                $WordCount = $WordList.Count
                $LineCount += $InputStorage.Count
                $PowerShell = [PowerShell]::Create().AddScript($InputProcessingScript).AddParameters(
                    @{
                        Strings       = $InputStorage.ToArray()
                        SplitChars    = $SplitChars
                        ExcludedWords = $StopWordsPattern
                    }
                )
                $PowerShell.RunspacePool = $RunspacePool

                $ProgressParams = @{
                    Activity         = "Processing Input Items: $LineCount"
                    Status           = "Jobs: <Run> $JobsStarted <Completed> $JobsReceived"
                    Id               = $ProgressID
                    CurrentOperation = "Splitting text into words. Word count: $WordCount"
                }
                Write-Progress @ProgressParams

                $RSJobs.Add(
                    [PSCustomObject]@{
                        Instance = $PowerShell
                        Result   = $PowerShell.BeginInvoke()
                    }
                )
                $JobsStarted ++
                $InputStorage.Clear()
            }

            do {
                $Completed = $RSJobs.Where{$_.Result.IsCompleted}
                foreach ($Runspace in $Completed) {
                    [string[]] $Result = $Runspace.Instance.EndInvoke($Runspace.Result).ForEach{$_}
                    $WordList.AddRange($Result)
                    $RSJobs.Remove($Runspace) > $null
                    $JobsReceived ++

                    $ProgressParams = @{
                        Activity         = "Processed Input Items: $LineCount"
                        Status           = "Jobs Run: $JobsStarted Completed: $JobsReceived"
                        Id               = $ProgressID
                        CurrentOperation = "Collating processed words from jobs"
                    }
                    Write-Progress @ProgressParams
                }

                Start-Sleep -Milliseconds 10
            } until ($RSJobs.Count -eq 0)
        }
        catch {
            $RunspacePool.Close()
            $RunspacePool.Dispose()
            $PSCmdlet.ThrowTerminatingError($_)
        }

        # Count occurrence of each word
        switch ($WordList.ToArray()) {
            { $WordHeightTable[($_ -replace 's$')] } {
                $WordHeightTable[($_ -replace 's$')] ++
                continue
            }
            { $WordHeightTable["${_}s"] } {
                $WordHeightTable[$_] = $WordHeightTable["${_}s"] + 1
                $WordHeightTable.Remove("${_}s")
                continue
            }
            default {
                $WordHeightTable[$_] ++
                continue
            }
        }

        if ($PSBoundParameters.ContainsKey('FocusWord')) {
            $WordHeightTable[$FocusWord] = $WordHeightTable.GetEnumerator() |
                ForEach-Object Value |
                Measure-Object -Maximum |
                ForEach-Object Maximum

            $WordHeightTable[$FocusWord] *= 1.25
        }

        $WordHeightTable | Out-String | Write-Debug

        $SortedWordList = $WordHeightTable.GetEnumerator().Name |
            Sort-Object -Descending { $WordHeightTable[$_] }

        if ($MaxUniqueWords) {
            $SortedWordList = $SortedWordList |
                Select-Object -First $MaxUniqueWords
        }

        $LowestFrequency, $HighestFrequency, $AverageFrequency = $SortedWordList |
            ForEach-Object { $WordHeightTable[$_] } |
            Measure-Object -Average -Maximum -Minimum |
            ForEach-Object {$_.Minimum, $_.Maximum, $_.Average}

        try {
            if ($BackgroundImage.FullName) {
                [Bitmap] $WordCloudImage = [Bitmap]::new($BackgroundImage.FullName)
                $WordCloudImage.SetResolution(96, 96)

                [Graphics] $DrawingSurface = [Graphics]::FromImage($WordCloudImage)
            }
            else {
                [Bitmap] $WordCloudImage = [Bitmap]::new($ImageSize.Width, $ImageSize.Height)
                $WordCloudImage.SetResolution(96, 96)

                [Graphics] $DrawingSurface = [Graphics]::FromImage($WordCloudImage)
                $DrawingSurface.Clear($BackgroundColor)
            }

            $DrawingSurface.PageScale = 1.0
            $DrawingSurface.PageUnit = $GraphicsUnit
            $DrawingSurface.SmoothingMode = [SmoothingMode]::AntiAlias
            $DrawingSurface.TextRenderingHint = [Text.TextRenderingHint]::ClearTypeGridFit

            $MaxSideLength = [Math]::Max($WordCloudImage.Width, $WordCloudImage.Height)

            Write-Verbose "Graphics Surface Properties:"
            $Properties = @(
                'RenderingOrigin'
                'CompositingQuality'
                'TextRenderingHint'
                'SmoothingMode'
                'PixelOffsetMode'
                'DpiX'
                'DpiY'
                'PageScale'
                'VisibleClipBounds'
            )
            $DrawingSurface |
                Select-Object -Property $Properties |
                Out-String |
                Write-Verbose

            Write-Verbose "Bitmap Properties:"
            $Properties = @(
                'Size'
                'PixelFormat'
                'RawFormat'
                'HorizontalResolution'
                'VerticalResolution'
            )
            $WordCloudImage |
                Select-Object -Property $Properties |
                Out-String |
                Write-Verbose

            $FontScale = $WordScale * 1.6 * ($WordCloudImage.Height + $WordCloudImage.Width) /
            ($AverageFrequency * $SortedWordList.Count)

            :size do {
                $ScaledWordHeightTable = @{}
                foreach ($Word in $SortedWordList) {
                    $ScaledWordHeightTable[$Word] = [Math]::Round(
                        $WordHeightTable[$Word] * $FontScale * (
                            2 * $RNG.NextDouble() /
                            (1 + $HighestFrequency - $LowestFrequency) + 0.9
                        )
                    )

                    if ($ScaledWordHeightTable[$Word] -lt 8) { continue }

                    $Font = [Font]::new(
                        $FontFamily,
                        $ScaledWordHeightTable[$Word],
                        $FontStyle,
                        $GraphicsUnit
                    )

                    $WordSize = $DrawingSurface.MeasureString($Word, $Font)
                    if ($Padding) {
                        $WordSize += [SizeF]::new($Padding, $Padding)
                    }
                    $WordFits = (
                        ($AllowOverflow -and [Math]::Max($WordSize.Width, $WordSize.Height)) -or
                        ($DisableWordRotation -and $WordSize.Width -lt $WordCloudImage.Width) -or
                        [Math]::Max($WordSize.Width, $WordSize.Height) -lt $MaxSideLength
                    )
                    if ($WordFits) {
                        $WordSizeTable[$Word] = $WordSize
                    }
                    else {
                        # Reset table and recalculate sizes (should only ever happen for the first few words at most)
                        $WordSizeTable.Clear()
                        $FontScale *= 0.98
                        continue size
                    }
                }

                # If we reach here, no words are larger than the image
                $SortedWordList = $SortedWordList.Where{$WordSizeTable.ContainsKey($_)}
                $MaxFontSize = $ScaledWordHeightTable[$SortedWordList[0]]
                $MinFontSize = $ScaledWordHeightTable[$Word]
                break
            } while ($true)
        }
        catch {
            $PSCmdlet.WriteError($_)
            if ($DrawingSurface) {
                $DrawingSurface.Dispose()
            }
            if ($DummyImage) {
                $WordCloudImage.Dispose()
            }
        }

        $GCD = Get-GreatestCommonDivisor -Numerator $MaxSideLength -Denominator ([Math]::Min($WordCloudImage.Width, $WordCloudImage.Height))
        $AspectRatio = $WordCloudImage.Width / $WordCloudImage.Height
        $CentrePoint = [PointF]::new($WordCloudImage.Width / 2, $WordCloudImage.Height / 2)

        [PSCustomObject]@{
            ExportFormat     = $ExportFormat
            TotalWords       = $WordList.Count
            UniqueWords      = $WordHeightTable.GetEnumerator().ForEach{$_}.Count
            DisplayedWords   = if ($MaxUniqueWords) {
                [Math]::Min($SortedWordList.Count, $MaxUniqueWords)
            }
            else {
                $SortedWordList
            }
            HighestFrequency = $HighestFrequency
            AverageFrequency = '{0:N2}' -f $AverageFrequency
            MaxFontSize      = $MaxFontSize
            MinFontSize      = $MinFontSize
            ImageSize        = $WordCloudImage.Size
            ImageCentre      = $CentrePoint
            AspectRatio      = "$($WordCloudImage.Width / $GCD) : $($WordCloudImage.Height / $GCD)"
            FontFamily       = $FontFamily.Name
        } | Out-String | Write-Verbose

        try {
            $BlankCanvas = $true

            [Region]$ForbiddenArea = [Region]::new()
            $ForbiddenArea.MakeInfinite()
            $UsableSpace = $DrawingSurface.VisibleClipBounds
            if ($AllowOverflow) {
                $UsableSpace.Inflate($UsableSpace.Width / 3, $UsableSpace.Width / 3)
            }

            $MaxRadialDistance = [Math]::Max($UsableSpace.Width, $UsableSpace.Height) / 2
            $ForbiddenArea.Exclude($UsableSpace)
            $WordCount = 0
            $WordPath = [GraphicsPath]::new()

            $RSJobs = [List[PSCustomObject]]::new()

            $RadialScanScript = {
                [CmdletBinding()]
                param(
                    $OccupiedSpace,
                    $Centre,
                    $WordSize,
                    $RadialAngles,
                    $RadialDistance,
                    $AspectRatio,
                    $DisableWordRotation,
                    $RandomSeed
                )

                $RNG = if ($null -ne $RandomSeed) {
                    [Random]::new($RandomSeed)
                }
                else {
                    [Random]::new()
                }

                # All type names in here use their full names; `using` declarations don't carry into new runspaces

                :angle foreach ($Angle in $RadialAngles) {
                    $Complex = [System.Numerics.Complex]::FromPolarCoordinates($RadialDistance, $Angle)
                    $OffsetX = $WordSize.Width * 0.5 * ($RNG.NextDouble() + 0.25)
                    $OffsetY = $WordSize.Height * 0.5 * ($RNG.NextDouble() + 0.25)

                    if ($DisableWordRotation) {
                        $FormatList = [System.Drawing.StringFormat]::new()
                    }
                    else {
                        $FormatList = @(
                            [System.Drawing.StringFormat]::new(),
                            [System.Drawing.StringFormat]::new([System.Drawing.StringFormatFlags]::DirectionVertical)
                        ) | Sort-Object { $RNG.Next() }
                    }

                    foreach ($Rotation in $FormatList) {
                        if ($Rotation.FormatFlags -eq [System.Drawing.StringFormatFlags]::DirectionVertical) {
                            $Point = [System.Drawing.PointF]::new(
                                $Complex.Real * $AspectRatio + $Centre.X - $OffsetY,
                                $Complex.Imaginary + $Centre.Y - $OffsetX
                            )

                            $RectangleSize = [System.Drawing.SizeF]::new($WordSize.Height, $WordSize.Width)
                            $Rectangle = [System.Drawing.RectangleF]::new($Point, $RectangleSize)
                        }
                        else {
                            $Point = [System.Drawing.PointF]::new(
                                $Complex.Real * $AspectRatio + $Centre.X - $OffsetX,
                                $Complex.Imaginary + $Centre.Y - $OffsetY
                            )

                            $Rectangle = [System.Drawing.RectangleF]::new($Point, $WordSize)
                        }


                        if (-not $OccupiedSpace.IsVisible($Rectangle)) {
                            [PSCustomObject]@{
                                Point    = $Point
                                Rotation = $Rotation
                            }

                            break angle
                        }
                    }
                }
            }

            :words foreach ($Word in $SortedWordList) {
                $WordCount ++
                $PlaceFound = $null
                $JobsStarted = 0
                $JobsReceived = 0

                $RadialDistance = 0
                $InflationValue = ( $ScaledWordHeightTable[$Word] / 10 ) * $Padding + $PenWidth

                :radialSearch do {
                    if (-not $BlankCanvas) {
                        $RadialDistance += (
                            $RNG.NextDouble() *
                            $ScaledWordHeightTable[$Word] *
                            $DistanceStep /
                            [Math]::Max(1, 21 - $Padding * 2)
                        )
                    }

                    $ProgressParams = @{
                        Activity         = "Generating word cloud"
                        Status           = "Searching for available space to draw '{0}' at {1} em" -f @(
                            $Word
                            $ScaledWordHeightTable[$Word]
                        )
                        CurrentOperation = "Paths: <Exhausted: {0}> <Searching: {1}>" -f $JobsReceived, $JobsStarted
                        PercentComplete  = ($WordCount / $SortedWordList.Count) * 100
                        Id               = $ProgressID
                    }
                    Write-Progress @ProgressParams

                    if ($RadialDistance -gt $MaxRadialDistance) {
                        continue words
                    }

                    $AngleIncrement = 360 / ( ($RadialDistance + 1) * $RadialGranularity / 10 )
                    $Start, $End = switch ([int]$RNG.Next() -band 7) {
                        0 { 0, 360 }
                        1 { -90, 270 }
                        2 { -180, 180 }
                        3 { -270, 90  }
                        4 { 360, 0; $AngleIncrement *= -1 }
                        5 { 270, -90; $AngleIncrement *= -1 }
                        6 { 180, -180; $AngleIncrement *= -1 }
                        7 { 90 - 270; $AngleIncrement *= -1 }
                    }

                    $Angle = $Start

                    $PowerShell = [powershell]::Create().AddScript($RadialScanScript).AddParameters(
                        @{
                            OccupiedSpace       = $ForbiddenArea.Clone()
                            Centre              = $CentrePoint
                            WordSize            = $WordSizeTable[$Word] + [SizeF]::new($InflationValue, $InflationValue)
                            RadialAngles        = do {
                                ConvertTo-Radians -Degrees $Angle

                                $Condition = if ($Start -lt $End ) {$Angle -le $End} else {$Angle -gt $End}
                                $Angle += $AngleIncrement
                            } while ($Condition)
                            RadialDistance      = $RadialDistance
                            AspectRatio         = $AspectRatio
                            DisableWordRotation = $DisableWordRotation
                            RandomSeed          = $RandomSeed
                        }
                    )
                    $PowerShell.RunspacePool = $RunspacePool

                    $RSJobs.Add(
                        [PSCustomObject]@{
                            Instance = $PowerShell
                            Result   = $PowerShell.BeginInvoke()
                        }
                    )
                    $JobsStarted ++

                    if ($BlankCanvas) {
                        Start-Sleep -Milliseconds 10
                    }

                    if ($BlankCanvas -or ($JobsStarted % 5 -eq 1) -or $RSJobs.Result.IsCompleted -contains $true) {
                        $FinishedJobList = $RSJobs | Where-Object { $_.Result.IsCompleted }

                        if (@($FinishedJobList).Count -gt 0) {
                            # We only care about the first available space
                            [PSObject]$Result = $FinishedJobList[0].Instance.EndInvoke($FinishedJobList[0].Result)

                            if ($Result.Count -gt 0) {
                                # There should only be one returned item in the collection
                                $PlaceFound = $Result[0]
                            }

                            # We really don't need any further results once we find one
                            $JobsReceived += $RSJobs.Count
                        }
                    }
                } until ($PlaceFound)

                $BlankCanvas = $false

                Write-Verbose "Location found for $Word"
                $WordPath.Reset()
                $WordPath.FillMode = [FillMode]::Winding
                $WordPath.AddString(
                    [string]$Word,
                    [FontFamily]$FontFamily,
                    [int]$FontStyle,
                    [float]$ScaledWordHeightTable[$Word],
                    [PointF]$PlaceFound.Point,
                    [StringFormat]$PlaceFound.Rotation
                )

                # Available location found; draw word and loop back to words
                if ($BoxCollisions) {
                    $Bounds = $WordPath.GetBounds()
                    $Bounds.Inflate($InflationValue, $InflationValue)
                    $ForbiddenArea.Union($Bounds)
                }
                else {
                    $ForbiddenArea.Union($WordPath)
                }

                $Color = $ColorList[$ColorIndex]
                $Brush = [SolidBrush]::new($Color)
                if ($StrokeWidth -gt 0) {
                    $PenWidth = $ScaledWordHeightTable[$Word] * ($StrokeWidth / 100)
                    $StrokePen = [Pen]::new([SolidBrush]::new($StrokeColor), $PenWidth)
                }

                $DrawingSurface.FillPath($Brush, $WordPath)
                if ($StrokePen) {
                    $DrawingSurface.DrawPath($StrokePen, $WordPath)
                }

                if ($BlankCanvas) {
                    $BlankCanvas = $false
                }

                $ColorIndex++
                if ($ColorIndex -ge $ColorList.Count) {
                    $ColorIndex = 0
                }

                # Make sure we reset our searcher pool for the next word
                foreach ($Job in $RSJobs) { $Job.Instance.EndInvoke($Job.Result) > $null }
                $RSJobs.Clear()
            }

            # All words written that we can, wait for any remaining draw operations to finish before saving.
            $DrawingSurface.Flush([FlushIntention]::Sync)

            foreach ($FilePath in $PathList) {
                if ($PSCmdlet.ShouldProcess($FilePath, 'Save word cloud to file')) {
                    $WordCloudImage.Save($FilePath, $ExportFormat)
                }
            }
        }
        catch {
            $PSCmdlet.WriteError($_)
        }
        finally {
            $RunspacePool.Dispose()
            $DrawingSurface.Dispose()
            $WordCloudImage.Dispose()
            $WordPath.Dispose()
        }
    }
}
