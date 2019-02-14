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
