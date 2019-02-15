namespace PSWordCloud
open System
open System.Collections.Generic
open System.Linq
open System.Management.Automation
open System.Numerics
open System.Text.RegularExpressions
open System.Threading.Tasks
open PSWordCloud.Extensions
open PSWordCloud.Utils
open SkiaSharp


[<Cmdlet(VerbsCommon.New, "WordCloud", DefaultParameterSetName = "ColorBackground")>]
[<Alias("nwc", "wcloud", "newcloud")>]
type NewWordCloudCommand() =
    inherit PSCmdlet()

    let mutable _resolvedPath = String.Empty
    let mutable _resolvedBackgroundPath = String.Empty
    let mutable _colors : SKColor list = []

    //#region Static Parameters
    static let _stopWords = [
        "a";"about";"above";"after";"again";"against";"all";"am";"an";"and";"any";"are";"aren't";"as";"at";"be";
        "because";"been";"before";"being";"below";"between";"both";"but";"by";"can't";"cannot";"could";"couldn't";
        "did";"didn't";"do";"does";"doesn't";"doing";"don't";"down";"during";"each";"few";"for";"from";"further";
        "had";"hadn't";"has";"hasn't";"have";"haven't";"having";"he";"he'd";"he'll";"he's";"her";"here";"here's";
        "hers";"herself";"him";"himself";"his";"how";"how's";"i";"i'd";"i'll";"i'm";"i've";"if";"in";"into";"is";
        "isn't";"it";"it's";"its";"itself";"let's";"me";"more";"most";"mustn't";"my";"myself";"no";"nor";"not";"of";
        "off";"on";"once";"only";"or";"other";"ought";"our";"ours";"ourselves";"out";"over";"own";"same";"shan't";
        "she";"she'd";"she'll";"she's";"should";"shouldn't";"so";"some";"such";"than";"that";"that's";"the";"their";
        "theirs";"them";"themselves";"then";"there";"there's";"these";"they";"they'd";"they'll";"they're";"they've";
        "this";"those";"through";"to";"too";"under";"until";"up";"very";"was";"wasn't";"we";"we'd";"we'll";"we're";
        "we've";"were";"weren't";"what";"what's";"when";"when's";"where";"where's";"which";"while";"who";"who's";
        "whom";"why";"why's";"with";"won't";"would";"wouldn't";"you";"you'd";"you'll";"you're";"you've";"your";
        "yours";"yourself";"yourselves"
    ]

    static let _splitChars = [
        ' ';'\n';'\t';'\r';'.';';';';';'\\';'/';'|';
        ':';'"';'?';'!';'{';'}';'[';']';':';'(';')';
        '<';'>';'“';'”';'*';'#';'%';'^';'&';'+';'='
    ]

    static let _randomLock = obj();
    static let mutable _random : Random = null
    static let random =
        if isNull _random then _random <- Random()
        _random
    //#endregion Static Parameters

    //#region Parameters

    [<Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "ColorBackground")>]
    [<Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "ColorBackground-Mono")>]
    [<Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "FileBackground")>]
    [<Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "FileBackground-Mono")>]
    [<Alias("InputString", "Text", "String", "Words", "Document", "Page")>]
    [<AllowEmptyString>]
    member val public InputObject : PSObject = null
        with get, set

    [<Parameter(Mandatory = true, Position = 0, ParameterSetName = "ColorBackground")>]
    [<Parameter(Mandatory = true, Position = 0, ParameterSetName = "ColorBackground-Mono")>]
    [<Parameter(Mandatory = true, Position = 0, ParameterSetName = "FileBackground")>]
    [<Parameter(Mandatory = true, Position = 0, ParameterSetName = "FileBackground-Mono")>]
    [<Alias("OutFile", "ExportPath", "ImagePath")>]
    member public self.Path
        with get() = _resolvedPath
        and set(v) = _resolvedPath <- self.SessionState.Path.GetUnresolvedProviderPathFromPSPath(v)

    [<Parameter(ParameterSetName = "ColorBackground")>]
    [<Parameter(ParameterSetName = "ColorBackground-Mono")>]
    [<ArgumentCompleter(typeof<SKSizeICompleter>)>]
    [<TransformToSKSizeI>]
    member val public ImageSize = SKSizeI(3840, 2160)
        with get, set

    [<Parameter(Mandatory = true, ParameterSetName = "FileBackground")>]
    [<Parameter(Mandatory = true, ParameterSetName = "FileBackground-Mono")>]
    member public self.BackgroundImage
        with get() = _resolvedBackgroundPath
        and set(v) =
            let previousDir = Environment.CurrentDirectory;
            Environment.CurrentDirectory <- self.SessionState.Path.CurrentFileSystemLocation.Path;
            _resolvedBackgroundPath <- System.IO.Path.GetFullPath(v);
            Environment.CurrentDirectory <- previousDir

    [<Parameter>]
    [<Alias("FontFamily", "FontFace")>]
    [<ArgumentCompleter(typeof<TypefaceCompleter>)>]
    [<TransformToSKTypeface()>]
    member val public Typeface = FontManager.MatchFamily("Consolas", SKFontStyle.Normal)
        with get, set

    [<Parameter(ParameterSetName = "ColorBackground")>]
    [<Parameter(ParameterSetName = "ColorBackground-Mono")>]
    [<Alias("Backdrop", "CanvasColor")>]
    [<ArgumentCompleter(typeof<SKColorCompleter>)>]
    [<TransformToSKColor>]
    member val public BackgroundColor = SKColors.Black
        with get, set

    [<Parameter>]
    [<SupportsWildcards>]
    [<TransformToSKColorAttribute>]
    [<ArgumentCompleter(typeof<SKColorCompleter>)>]
    member val public ColorSet = StandardColors |> Seq.toList
        with get, set

    [<Parameter>]
    [<Alias("OutlineWidth")>]
    [<ValidateRange(0, 10)>]
    member val public StrokeWidth = 0.0f
        with get, set

    [<Parameter>]
    [<Alias("OutlineColor")>]
    [<TransformToSKColor>]
    [<ArgumentCompleter(typeof<SKColorCompleter>)>]
    member val public StrokeColor = SKColors.Black
        with get, set

    [<Parameter>]
    [<Alias("Title")>]
    member val public FocusWord = String.Empty
        with get, set

    [<Parameter>]
    [<Alias("ScaleFactor")>]
    [<ValidateRange(0.01, 20)>]
    member val public WordScale = 1.0f
        with get, set

    [<Parameter>]
    [<Alias("Spacing")>]
    member val public Padding = 3.0f
        with get, set

    [<Parameter>]
    [<ValidateRange(1, 500)>]
    member val public DistanceStep = 5.0f
        with get, set

    [<Parameter>]
    [<ValidateRange(1, 50)>]
    member val public RadialStep = 15.0f
        with get, set

    [<Parameter>]
    [<Alias("MaxWords")>]
    [<ValidateRange(0, Int32.MaxValue)>]
    member val public MaxRenderedWords = 100
        with get, set

    [<Parameter>]
    [<Alias("MaxColours")>]
    [<ValidateRange(1, Int32.MaxValue)>]
    member val public MaxColors = Int32.MaxValue
        with get, set

    [<Parameter>]
    [<Alias("SeedValue")>]
    member val public RandomSeed = 0
        with get, set

    [<Parameter>]
    [<Alias("DisableWordRotation")>]
    member val public DisableRotation : SwitchParameter = SwitchParameter(false)
        with get, set

    [<Parameter(Mandatory = true, ParameterSetName = "FileBackground-Mono")>]
    [<Parameter(Mandatory = true, ParameterSetName = "ColorBackground-Mono")>]
    [<Alias("Greyscale")>]
    member val public Monochrome : SwitchParameter = SwitchParameter(false)
        with get, set

    [<Parameter>]
    [<Alias("IgnoreStopWords")>]
    member val public AllowStopWords : SwitchParameter = SwitchParameter(false)
        with get, set

    [<Parameter>]
    member val public PassThru : SwitchParameter = SwitchParameter(false)
        with get, set

    [<Parameter>]
    [<Alias("AllowBleed")>]
    member val public AllowOverflow : SwitchParameter = SwitchParameter(false)
        with get, set

    //#endregion Parameters