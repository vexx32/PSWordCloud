---
external help file: PSWordCloudCmdlet.dll-Help.xml
Module Name: PSWordCloud
online version: 2.0
schema: 2.0.0
---

# New-WordCloud

## SYNOPSIS

Describes the syntax and behaviour of the New-WordCloud cmdlet.

## SYNTAX

### ColorBackground (Default)
```
New-WordCloud -InputObject <PSObject> [-Path] <String> [-ImageSize <SKSizeI>] [-Typeface <SKTypeface>]
 [-BackgroundColor <SKColor>] [-ColorSet <SKColor[]>] [-StrokeWidth <Single>] [-StrokeColor <SKColor>]
 [-ExcludeWord <String[]>] [-IncludeWord <String[]>] [-WordScale <Single>] [-AllowRotation <WordOrientations>]
 [-Padding <Single>] [-DistanceStep <Single>] [-RadialStep <Single>] [-MaxRenderedWords <Int32>]
 [-MaxColors <Int32>] [-RandomSeed <Int32>] [-Monochrome] [-AllowStopWords] [-AllowOverflow] [-PassThru]
 [<CommonParameters>]
```

### ColorBackground-FocusWord
```
New-WordCloud -InputObject <PSObject> [-Path] <String> [-ImageSize <SKSizeI>] [-Typeface <SKTypeface>]
 [-BackgroundColor <SKColor>] [-ColorSet <SKColor[]>] [-StrokeWidth <Single>] [-StrokeColor <SKColor>]
 -FocusWord <String> [-RotateFocusWord <Single>] [-ExcludeWord <String[]>] [-IncludeWord <String[]>]
 [-WordScale <Single>] [-AllowRotation <WordOrientations>] [-Padding <Single>] [-DistanceStep <Single>]
 [-RadialStep <Single>] [-MaxRenderedWords <Int32>] [-MaxColors <Int32>] [-RandomSeed <Int32>] [-Monochrome]
 [-AllowStopWords] [-AllowOverflow] [-PassThru] [<CommonParameters>]
```

### FileBackground
```
New-WordCloud -InputObject <PSObject> [-Path] <String> -BackgroundImage <String> [-Typeface <SKTypeface>]
 [-ColorSet <SKColor[]>] [-StrokeWidth <Single>] [-StrokeColor <SKColor>] [-ExcludeWord <String[]>]
 [-IncludeWord <String[]>] [-WordScale <Single>] [-AllowRotation <WordOrientations>] [-Padding <Single>]
 [-DistanceStep <Single>] [-RadialStep <Single>] [-MaxRenderedWords <Int32>] [-MaxColors <Int32>]
 [-RandomSeed <Int32>] [-Monochrome] [-AllowStopWords] [-AllowOverflow] [-PassThru] [<CommonParameters>]
```

### FileBackground-FocusWord
```
New-WordCloud -InputObject <PSObject> [-Path] <String> -BackgroundImage <String> [-Typeface <SKTypeface>]
 [-ColorSet <SKColor[]>] [-StrokeWidth <Single>] [-StrokeColor <SKColor>] -FocusWord <String>
 [-RotateFocusWord <Single>] [-ExcludeWord <String[]>] [-IncludeWord <String[]>] [-WordScale <Single>]
 [-AllowRotation <WordOrientations>] [-Padding <Single>] [-DistanceStep <Single>] [-RadialStep <Single>]
 [-MaxRenderedWords <Int32>] [-MaxColors <Int32>] [-RandomSeed <Int32>] [-Monochrome] [-AllowStopWords]
 [-AllowOverflow] [-PassThru] [<CommonParameters>]
```

### ColorBackground-WordTable
```
New-WordCloud -WordSizes <IDictionary> [-Path] <String> [-ImageSize <SKSizeI>] [-Typeface <SKTypeface>]
 [-BackgroundColor <SKColor>] [-ColorSet <SKColor[]>] [-StrokeWidth <Single>] [-StrokeColor <SKColor>]
 [-ExcludeWord <String[]>] [-IncludeWord <String[]>] [-WordScale <Single>] [-AllowRotation <WordOrientations>]
 [-Padding <Single>] [-DistanceStep <Single>] [-RadialStep <Single>] [-MaxRenderedWords <Int32>]
 [-MaxColors <Int32>] [-RandomSeed <Int32>] [-Monochrome] [-AllowStopWords] [-AllowOverflow] [-PassThru]
 [<CommonParameters>]
```

### ColorBackground-FocusWord-WordTable
```
New-WordCloud -WordSizes <IDictionary> [-Path] <String> [-ImageSize <SKSizeI>] [-Typeface <SKTypeface>]
 [-BackgroundColor <SKColor>] [-ColorSet <SKColor[]>] [-StrokeWidth <Single>] [-StrokeColor <SKColor>]
 -FocusWord <String> [-RotateFocusWord <Single>] [-ExcludeWord <String[]>] [-IncludeWord <String[]>]
 [-WordScale <Single>] [-AllowRotation <WordOrientations>] [-Padding <Single>] [-DistanceStep <Single>]
 [-RadialStep <Single>] [-MaxRenderedWords <Int32>] [-MaxColors <Int32>] [-RandomSeed <Int32>] [-Monochrome]
 [-AllowStopWords] [-AllowOverflow] [-PassThru] [<CommonParameters>]
```

### FileBackground-WordTable
```
New-WordCloud -WordSizes <IDictionary> [-Path] <String> -BackgroundImage <String> [-Typeface <SKTypeface>]
 [-ColorSet <SKColor[]>] [-StrokeWidth <Single>] [-StrokeColor <SKColor>] [-ExcludeWord <String[]>]
 [-IncludeWord <String[]>] [-WordScale <Single>] [-AllowRotation <WordOrientations>] [-Padding <Single>]
 [-DistanceStep <Single>] [-RadialStep <Single>] [-MaxRenderedWords <Int32>] [-MaxColors <Int32>]
 [-RandomSeed <Int32>] [-Monochrome] [-AllowStopWords] [-AllowOverflow] [-PassThru] [<CommonParameters>]
```

### FileBackground-FocusWord-WordTable
```
New-WordCloud -WordSizes <IDictionary> [-Path] <String> -BackgroundImage <String> [-Typeface <SKTypeface>]
 [-ColorSet <SKColor[]>] [-StrokeWidth <Single>] [-StrokeColor <SKColor>] -FocusWord <String>
 [-RotateFocusWord <Single>] [-ExcludeWord <String[]>] [-IncludeWord <String[]>] [-WordScale <Single>]
 [-AllowRotation <WordOrientations>] [-Padding <Single>] [-DistanceStep <Single>] [-RadialStep <Single>]
 [-MaxRenderedWords <Int32>] [-MaxColors <Int32>] [-RandomSeed <Int32>] [-Monochrome] [-AllowStopWords]
 [-AllowOverflow] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION

New-WordCloud takes input text either over the pipeline or directly to its -InputObject parameter,
and uses the word frequency distribution to create a word cloud image. More frequently-used words
are rendered larger in the final image.

## EXAMPLES

### Example 1

```powershell
PS> Get-Content .\MyEntireBook.txt | New-WordCloud -Path .\BookCloud.svg
```

Creates a word cloud and saves it to a file called BookCloud.svg in the current folder.

The default font, color, and layout settings will be used.

### Example 2

```powershell
PS> Get-ChildItem .\scripts -Recurse -Filter "*.ps1" |
>> Get-Content |
>> New-WordCloud -Path .\BookCloud.svg -FocusWord Scripts -ColorSet *blue*, *white*, *tan*, *yellow*, *gold* -Typeface Scriptina -StrokeWidth 2 -StrokeColor Brown -MaxRenderedWords 250
```

Creates a word cloud from all PS1 files in the scripts directory, with the word "Scripts" emblazoned
in the centre.

Word colors will be chosen from the specified set, including all named SKColors that match the
specified patterns.

The font Scriptina will be used, and all words will have a brown stroke (outline).

Up to 250 words will be used in the final image.

### Example 3

```powershell
PS> $Params = @{
    Path        = '.\BookCloud.svg'
    FocusWord   = 'News'
    ColorSet    = @(
        @{Red = 200; Green = 180; Blue = 54}
        "1899FF"
        "*white*"
        "*gold*"
    )
    StrokeWidth = 0
    StrokeColor = @{Red = 10; Green = 5; Blue = 45; Alpha = 128}
}
PS> Get-Content .\Newspaper.md | New-WordCloud @Params
```

Creates a word cloud from the Newspaper.md file in the current directory, with the word "News"
emblazoned in the centre.

Word colors will be chosen from the specified set, including all named SKColors that match the
specified patterns and those specified by use of the hashtable and hex-string values.

The default font will be used, and all words will have a hairline-width semi-transparent dark blue
stroke (outline).

## PARAMETERS

### -AllowOverflow

This option permits the word cloud to overflow the bounds of the canvas.

Use this in conjunction with the -WordScale parameter to create partially-clipped word clouds.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: AllowBleed

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AllowRotation

Specify -AllowRotation None to prevent word rotation entirely, or use one of the options to permit specific rotation modes.

All modes permit the "upright" standard orientation, as well as their specified additions.

```yaml
Type: WordOrientations
Parameter Sets: (All)
Aliases:
Accepted values: None, Vertical, FlippedVertical, EitherVertical, UprightDiagonals, InvertedDiagonals, AllDiagonals, AllUpright, AllInverted, All

Required: False
Position: Named
Default value: EitherVertical
Accept pipeline input: False
Accept wildcard characters: False
```

### -AllowStopWords

This option disables the use of the standard Stop Words list. By default, the following words are
ignored during text processing as they otherwise typically are extremely common and would dominate
the word cloud:

a, about, above, after, again, against, all, am, an, and, any, are, aren't, as, at, be, because,
been, before, being, below, between, both, but, by, can't, cannot, could, couldn't, did, didn't,
do, does, doesn't, doing, don't, down, during, each, few, for, from, further, had, hadn't, has,
hasn't, have, haven't, having, he, he'd, he'll, he's, her, here, here's, hers, herself, him,
himself, his, how, how's, I, I'd, I'll, I'm, I've, if, in, into, is, isn't, it, it's, its,
itself, let's, me, more, most, mustn't, my, myself, no, nor, not, of, off, on, once, only, or,
other, ought, our, ours, ourselves, out, over, own, same, shan't, she, she'd, she'll, she's, should,
shouldn't, so, some, such, than, that, that's, the, their, theirs, them, themselves, then, there,
there's, these, they, they'd, they'll, they're, they've, this, those, through, to, too, under,
until, up, very, was, wasn't, we, we'd, we'll, we're, we've, were, weren't, what, what's, when,
when's, where, where's, which, while, who, who's, whom, why, why's, with, won't, would, wouldn't,
you, you'd, you'll, you're, you've, your, yours, yourself, yourselves

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: IgnoreStopWords

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BackgroundColor

Specifies the SKColor values used for the background of the canvas. Multiple values are accepted,
and each new word pulls the next color from the set. If the end of the set is reached, the next word
will reset the index to the start and retrieve the first color again.

Accepts input as a complete SKColor object, or one of the following formats:

1. One or more strings matching one of the named color fields in [SkiaSharp.SKColors]. These values
   will be pulled for tab-completion automatically. Names containing wildcards may be used, and all
   matching colors will be included in the set. The value "Transparent" is also accepted here.
2. A hexadecimal number string with or without the preceding #, in the form: AARRGGBB, RRGGBB, ARGB,
   or RGB.
3. A hashtable or custom object with keys or properties named: "Red, Green, Blue", and/or "Alpha",
   with values may range from 0-255. Omitted color values are assumed to be 0, but omitting alpha
   defaults it to 255 (fully opaque).

```yaml
Type: SKColor
Parameter Sets: ColorBackground, ColorBackground-FocusWord, ColorBackground-WordTable, ColorBackground-FocusWord-WordTable
Aliases: Backdrop, CanvasColor

Required: False
Position: Named
Default value: Black
Accept pipeline input: False
Accept wildcard characters: False
```

### -BackgroundImage

Specifies the path to the background image to be used as a base for the final word cloud image.

```yaml
Type: String
Parameter Sets: FileBackground, FileBackground-FocusWord, FileBackground-WordTable, FileBackground-FocusWord-WordTable
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ColorSet

Specifies the SKColor values used for the words in the cloud. Multiple values are accepted, and each new word pulls the next color from the set. If the end of the set is reached, the next word will reset the index to the start and retrieve the first color again.

Accepts input as a complete SKColor object, or one of the following formats:

1. One or more strings matching one of the named color fields in [SkiaSharp.SKColors]. These values will be pulled for tab-completion automatically. Names containing wildcards may be used, and all matching colors will be included in the set.
2. A hexadecimal number string with or without the preceding #, in the form: AARRGGBB, RRGGBB, ARGB, or RGB.
3. A hashtable or custom object with keys or properties named: "Red, Green, Blue", and/or "Alpha", with values may range from 0-255. Omitted color values are assumed to be 0, but omitting alpha defaults it to 255 (fully opaque).

```yaml
Type: SKColor[]
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: *
Accept pipeline input: False
Accept wildcard characters: False
```

### -DistanceStep

Determines the value to scale the distance step by.
Larger numbers will result in more radially spread-out clouds.

```yaml
Type: Single
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: 5.0
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExcludeWord

Determines the words to be explicitly ignored when rendering the word cloud.
This is usually used to exclude irrelevant words, unwanted URL segments, etc.

Values from -IncludeWord take precedence over those from this parameter.

```yaml
Type: String[]
Parameter Sets: (All)
Aliases: ForbidWord, IgnoreWord

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FocusWord

Determines the focus word string to be used in the word cloud. This string will typically appear in
the centre of the cloud, larger than all the other words.

```yaml
Type: String
Parameter Sets: ColorBackground-FocusWord, FileBackground-FocusWord, ColorBackground-FocusWord-WordTable, FileBackground-FocusWord-WordTable
Aliases: Title

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ImageSize

Determines the canvas size for the word cloud image.

Input can be passed directly as a [SkiaSharp.SKSizeI] object, or in one of the following formats:

1. A predefined size string. One of:
     - 720p           (canvas size: 1280x720)
     - 1080p          (canvas size: 1920x1080)
     - 4K             (canvas size: 3840x2160)
     - A4             (canvas size: 816x1056)
     - Poster11x17    (canvas size: 1056x1632)
     - Poster18x24    (canvas size: 1728x2304)
     - Poster24x36    (canvas size: 2304x3456)

2. Single integer (e.g., -ImageSize 1024). This will be used as both the width and height of the image, creating a square canvas.
3. Any image size string (e.g., 1024x768). The first number will be used as the width, and the second number used as the height of the canvas.
4. A hashtable or custom object with keys or properties named "Width" and "Height" that contain integer values

```yaml
Type: SKSizeI
Parameter Sets: ColorBackground, ColorBackground-FocusWord, ColorBackground-WordTable, ColorBackground-FocusWord-WordTable
Aliases:

Required: False
Position: Named
Default value: 3840x2160
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludeWord

Specifies normally-excluded words to include in the word cloud.

This parameter takes precedence over the -ExcludeWord parameter.

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -InputObject

Provides the input text to supply to the word cloud.
All input is accepted, but will be treated as string data regardless of the input type.
If you are entering complex object input, ensure the objects have a meaningful ToString() method override defined.

```yaml
Type: PSObject
Parameter Sets: ColorBackground, ColorBackground-FocusWord, FileBackground, FileBackground-FocusWord
Aliases: InputString, Text, String, Words, Document, Page

Required: True
Position: Named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -MaxColors

Determines the maximum number of colors to use from the values contained in the -ColorSet parameter.

The values from the -ColorSet parameter are shuffled before being trimmed down here, so that you are given a variety of color selections even under default conditions.

```yaml
Type: Int32
Parameter Sets: (All)
Aliases: MaxColours

Required: False
Position: Named
Default value: [int]::MaxValue
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxRenderedWords

Specifies the maximum number of words to draw in the rendered cloud.
More words take longer to render, and after a few hundred words the visible sizes become unreadable in all but the largest images.

However, an appropriate vector graphics viewer or editor is still capable of zooming in far enough to see them.

```yaml
Type: Int32
Parameter Sets: (All)
Aliases: MaxWords

Required: False
Position: Named
Default value: 100
Accept pipeline input: False
Accept wildcard characters: False
```

### -Monochrome

If this option is specified, New-WordCloud draws the word cloud in monochrome (greyscale).
Only the Brightness values from the SKColors in the color set provided will be used.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: BlackAndWhite, Greyscale

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Padding

Determines the float value to scale the padding space around the words by.

```yaml
Type: Single
Parameter Sets: (All)
Aliases: Spacing

Required: False
Position: Named
Default value: 5.0
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru

Specifying this option causes New-WordCloud to emit a FileInfo object representing the finished SVG file when it is completed.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path

The output file path to save the final SVG vector image to. Output is written as a stream, while the image is being generated. Terminating the command early may result in a usable but partially-formed image, or an invalid SVG file with missing tags.

```yaml
Type: String
Parameter Sets: (All)
Aliases: OutFile, ExportPath, ImagePath

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RadialStep

Determines the distance around each radial arc that the scanning algorithm takes for each circular sweep.
Larger values correspond to fewer points checked on each radial sweep.
This value is scaled according to distance from the center, so there will be more steps on a larger radius.

```yaml
Type: Single
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: 15.0
Accept pipeline input: False
Accept wildcard characters: False
```

### -RandomSeed

Determines the seed value for the random numbers used to vary the position and placement patterns.

```yaml
Type: Int32
Parameter Sets: (All)
Aliases: SeedValue

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RotateFocusWord

Specify an angle in degrees to rotate the focus word by, overriding the default random rotations for the focus word only.

Values from -360 to 360, including sub-degree increments, are permitted.

```yaml
Type: Single
Parameter Sets: ColorBackground-FocusWord, FileBackground-FocusWord, ColorBackground-FocusWord-WordTable, FileBackground-FocusWord-WordTable
Aliases: RotateTitle

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -StrokeColor

Determines the SKColor value used as the stroke color for the words in the image.
Accepts input as a complete SKColor object, or one of the following formats:

1. A string color name matching one of the fields in SkiaSharp.SKColors.
These values will be pulled for tab-completion automatically.
Wildcards may be used only if the pattern matches exactly one color name.
2. A hexadecimal number string with or without the preceding #, in the form: AARRGGBB, RRGGBB, ARGB, or RGB.
3. A hashtable or custom object with keys or properties named: "Red, Green, Blue", and/or "Alpha", with values from 0-255.
Omitted color values are assumed to be 0, but omitting alpha defaults it to 255 (fully opaque)

```yaml
Type: SKColor
Parameter Sets: (All)
Aliases: OutlineColor

Required: False
Position: Named
Default value: Black
Accept pipeline input: False
Accept wildcard characters: False
```

### -StrokeWidth

Determines the width of the word outline.
Values from 0-10 are  permitted.
A zero value indicates the special "Hairline" width, where the width of the stroke depends on the SVG viewing scale.

```yaml
Type: Single
Parameter Sets: (All)
Aliases: OutlineWidth

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Typeface

Gets or sets the typeface to be used in the word cloud.
Input can be processed as a SkiaSharp.SKTypeface object, or one of the following formats:

1. String value matching a valid font name. These can be autocompleted by pressing [Tab].
An invalid value will cause the system default to be used.
2. A custom object or hashtable object containing the following keys or properties:
    - FamilyName: string value. If no font by this name is available, the system default will be used.
    - FontWeight: "Invisible", "ExtraLight", Light", "Thin", "Normal", "Medium", "SemiBold", "Bold",
      "ExtraBold", "Black", "ExtraBlack" (Default: "Normal")
    - FontSlant: "Upright", "Italic", "Oblique" (Default: "Upright")
    - FontWidth: "UltraCondensed", "ExtraCondensed", "Condensed", "SemiCondensed", "Normal", "SemiExpanded",
      "Expanded", "ExtraExpanded", "UltraExpanded" (Default: "Normal")

```yaml
Type: SKTypeface
Parameter Sets: (All)
Aliases: FontFamily, FontFace

Required: False
Position: Named
Default value: Consolas
Accept pipeline input: False
Accept wildcard characters: False
```

### -WordScale

Applies a scaling value to the words in the cloud.
Use this parameter to shrink or expand your total word cloud area with respect to the size of the image.
The default of 1.0 is approximately equivalent to the total image size.
Scale as appropriate according to how much of the total canvas you would like the cloud to cover.

The cloud size is restricted to the canvas size by default, so values above 1.0 will typically not have an impact without also supplying the -AllowOverflow option.

```yaml
Type: Single
Parameter Sets: (All)
Aliases: ScaleFactor

Required: False
Position: Named
Default value: 1.0
Accept pipeline input: False
Accept wildcard characters: False
```

### -WordSizes

Instead of supplying a chunk of text as the input, this parameter allows you to define your own relative word sizes.
Supply a dictionary or hashtable object where the keys are the words you want to draw in the cloud, and the values are their relative sizes.
Words will be scaled as a percentage of the largest sized word in the table.
In other words, if you have @{ text = 10; image = 100 }, then "text" will appear 10 times smaller than "image".

```yaml
Type: IDictionary
Parameter Sets: ColorBackground-WordTable, ColorBackground-FocusWord-WordTable, FileBackground-WordTable, FileBackground-FocusWord-WordTable
Aliases: WordSizeTable, CustomWordSizes

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

### System.Management.Automation.PSObject

New-WordCloud accepts pipeline input of any type to its -InputObject parameter.
Due to the nature of the command, all inputs will be transformed to string before they are used in the final word cloud.
Complex objects may be reduced to their type names only, if they do not have a predefined conversion path to a string representation.

## OUTPUTS

### System.IO.FileInfo

If the -PassThru switch is used, New-WordCloud will output the FileInfo object representing the completed image file.
Otherwise, there is no output to the console.

## NOTES

Due to its dependence on the SkiaSharp library, loading the New-WordCloud module will also expose the SkiaSharp library types for you to use.
This is both by necessity and for configurability.
SkiaSharp types are accessible in the SkiaSharp namespace, for example [SkiaSharp.SKTypeface].
It is also possible to surface the type names with a using namespace declaration.

While a lot of work has gone into the parameter transforms to ensure you can customise the final look of the word cloud as much as possible, New-WordCloud seamlessly accepts direct input of the SkiaSharp objects it utilises, so that you can obtain and use your own SkiaSharp library objects for maximum configurability.

## RELATED LINKS

[Online Version](https://github.com/vexx32/PSWordCloud/blob/master/docs/New-WordCloud.md)

[SkiaSharp API Reference](https://docs.microsoft.com/en-us/dotnet/api/skiasharp)
