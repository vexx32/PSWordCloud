# PSWordCloud Examples

Below are some sample output images, along with as many of their input commands as I can recall.
Some of these may be from slightly different versions of the module, but by and large they should represent fairly accurately the results you can obtain with a little exploration.

Note that some of these are rendered on transparent backgrounds and may not look the best on lighter backdrops.

## v2.x.x Examples

### New Rotation Modes Tests

The below examples use the following command to generate the cloud, only adding the -AllowRotation mode for that test:

```powershell
$poem | New-WordCloud -Path .\test.svg -ImageSize 800 -Typeface "Segoe Print" -FocusWord Hollow
```

`$poem` contains the poem text from [this page](http://www.poetsforum.com/poems/hollow/).

#### -AllowRotation None

![None](./_images/NoRotation.svg)

#### -AllowRotation Vertical

![Vertical](./_images/VerticalRotationR.svg)

#### -AllowRotation FlippedVertical

![FlippedVertical](./_images/VerticalRotationL.svg)

#### -AllowRotation EitherVertical (Default)

![EitherVertical](./_images/VerticalRotation2.svg)

#### -AllowRotation UprightDiagonals

![UprightDiagonals](./_images/UprightDiag.svg)

#### -AllowRotation InvertedDiagonals

![InvertedDiagonals](./_images/InvertedDiag.svg)

#### -AllowRotation AllDiagonals

![AllDiagonals](./_images/AllDiag.svg)

#### -AllowRotation AllUpright

![AllUpright](./_images/AllUpright.svg)

#### -AllowRotation AllInverted

![AllInverted](./_images/AllInverted.svg)

#### -AllowRotation All

![All](./_images/AllRotations.svg)

## v1.x.x Examples

These examples come from the v1.x.x versions of PSWordCloud.

### Edge of Night

Input is the complete Lyrics from _Edge of Night_, the song Pipping sings in _Lord of the Rings_ to the Steward of Gondor.

![Edge of Night](./_images/EdgeOfNight.png)

```powershell
$Params = @{
    Path            = '.\EdgeOfNight.png'
    FocusWord       = 'Edge of Night'
    StrokeWidth     = 2
    StrokeColor     = 'MidnightBlue'
    FontFamily      = 'Hobbiton Brushhand'
    ImageSize       = 4096
    BackgroundColor = 'Transparent'
    Padding         = 5
}
$Lyrics | New-WordCloud @Params
```

### Paterson

These are a collection I created for a friend of my wife and I. Here's to William Carlos Williams' _Paterson_, just for you, Kate!

#### Cursive

![Paterson Cursive](./_images/Paterson-Cursive.png)

This one was very standard settings, apart from the font, which I don't recall.

#### Scriptina

![Paterson Scriptina Blue](./_images/Paterson-Scriptina.png)

```powershell
$Params = @{
    Path            = '.\Paterson.png'
    FocusWord       = 'Paterson'
    StrokeWidth     = 1
    StrokeColor     = 'Blue'
    FontFamily      = 'Scriptina'
    ImageSize       = 4096
    BackgroundColor = 'Transparent'
}
$Paterson | New-WordCloud @Params
```

![Paterson Scriptina Deep Blue](./_images/Paterson-Scriptina2.png)

```powershell
$Params = @{
    Path            = '.\Paterson.png'
    FocusWord       = 'Paterson'
    StrokeWidth     = 1
    StrokeColor     = 'MidnightBlue'
    FontFamily      = 'Scriptina'
    ImageSize       = 4096
    BackgroundColor = 'Transparent'
}
$Paterson | New-WordCloud @Params
```

![Paterson Scriptina Brown](./_images/Paterson-ScriptinaBrown.png)

```powershell
$Params = @{
    Path            = '.\Paterson.png'
    FocusWord       = 'Paterson'
    StrokeWidth     = 1
    StrokeColor     = 'Brown'
    FontFamily      = 'Scriptina'
    ImageSize       = 4096
    BackgroundColor = 'Transparent'
}
$Paterson | New-WordCloud @Params
```

![Paterson Scriptina Alt](./_images/Paterson-Scriptina3.png)

```powershell
$Params = @{
    Path            = '.\Paterson.png'
    FocusWord       = 'Paterson'
    StrokeWidth     = 1
    StrokeColor     = 'MidnightBlue'
    ColorSet        = '*light*'
    FontFamily      = 'Scriptina'
    ImageSize       = 4096
    BackgroundColor = 'Transparent'
}
$Paterson | New-WordCloud @Params
```

### PSKoans

These are all created from the text and script data in the PSKoans module.
Some include the test files, but most do not.

![PSKoans HBH](./_images/PSKoans_HBH.png)

```powershell
$Params = @{
    Path            = '.\PSKoans.png'
    FocusWord       = 'PSKoans'
    FontFamily      = 'Hobbiton Brushhand'
    ColorSet        = '*light*'
    ImageSize       = 4096
    BackgroundColor = 'Transparent'
}
$PSKoans | New-WordCloud @Params
```

![PSKoans HBH](./_images/PSKoans-motivational.png)

```powershell
$Params = @{
    Path            = '.\PSKoans.png'
    FontFamily      = 'Nerwus'
    StrokeWidth     = 1
    StrokeColor     = 'Brown'
    BackgroundColor = 'Transparent'
}
$PSKoans | New-WordCloud @Params
```

![PSKoans Pretty](./_images/PSKoans-pretty.png)

```powershell
$Params = @{
    Path            = '.\PSKoans.png'
    FontFamily      = 'Scriptina'
    ImageSize       = 4096
    StrokeWidth     = 1
    StrokeColor     = 'Brown'
    BackgroundColor = 'Transparent'
}
$PSKoans | New-WordCloud @Params
```

![PSKoans Zombie](./_images/PSKoans-zombieholocaust.png)

```powershell
$Params = @{
    Path            = '.\PSKoans.png'
    FontFamily      = 'Zombie Holocaust'
    ImageSize       = 4096
    StrokeWidth     = 1
    StrokeColor     = 'Brown'
    BackgroundColor = 'Transparent'
}
$PSKoans | New-WordCloud @Params
```

![PSKoans Wingdings](./_images/Wingdings.png)

```powershell
$Params = @{
    Path            = '.\PSKoans-Wingdings.png'
    ImageSize       = 4096
    FontFamily      = 'Wingdings'
    BackgroundColor = 'Transparent'
}
$PSKoans | New-WordCloud @Params
```

### PowerShell

These were created using _every line of code_ from every `.cs` file in the [PowerShell Core](https://github.com/PowerShell/PowerShell) GitHub repository.

This was so slow I was literally forced to apply multithreading to the module to make it complete _just the text processing portion_ in under an hour on my machine.
It's much quicker now, but these were each an ordeal.

![PowerShell](./_images/PowerShell-one.png)

```powershell
$Params = @{
    Path            = '.\PowerShell.png'
    FocusWord       = 'PowerShell'
    StrokeWidth     = 1
    StrokeColor     = 'MidnightBlue'
    ImageSize       = 4096
    BackgroundColor = 'Transparent'
    Padding         = 2
    MaxUniqueWords  = 250
    ExcludeWord     = Get-Content -Path .\ListofCommonCSharpKeywords.txt
}
$PSCoreCode | New-WordCloud @Params
```

![PowerShell](./_images/PowerShell-Title.png)

```powershell
$Params = @{
    Path            = '.\PowerShell.png'
    FocusWord       = 'PowerShell'
    StrokeWidth     = 1
    StrokeColor     = 'MidnightBlue'
    ColorSet        = '*light*', '*blue*', '*purple*'
    ImageSize       = 4096
    BackgroundColor = 'Transparent'
    Padding         = 2
}
$PSCoreCode | New-WordCloud @Params
```

![PowerShell](./_images/PSCore-black.png)

```powershell
$Params = @{
    Path            = '.\PowerShell.png'
    MaxUniqueWords  = 350
    Padding         = 2
}
$PSCoreCode | New-WordCloud @Params
```

![PowerShell](./_images/PSCore-White.png)

```powershell
$Params = @{
    Path            = '.\PowerShell.png'
    MaxUniqueWords  = 0
    Padding         = 0
    StrokeWidth     = 1
    StrokeColor     = 'MidnightBlue'
    ColorSet        = '*blue*','*purple*'
    BackgroundColor = 'White'
}
$PSCoreCode | New-WordCloud @Params
```

### PowerShell-RFC Repo

This one is made from all the [PowerShell-RFC](https://github.com/PowerShell/PowerShell-RFC) Markdown files.

![PowerShell-RFC](./_images/PowerShellRFC.png)

```powershell
$Params = @{
    Path            = '.\PowerShell-RFC.png'
    Padding         = 3
    StrokeWidth     = 1
    StrokeColor     = 'MidnightBlue'
    ColorSet        = '*blue*','*purple*'
    BackgroundColor = 'White'
    FocusWord       = 'RFC'
}
$PSRFCMarkdown | New-WordCloud @Params
```
