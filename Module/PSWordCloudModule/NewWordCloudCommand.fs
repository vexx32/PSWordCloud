namespace PSWordCloud
open System
open System.Collections.Generic
open System.Linq
open System.Management.Automation
open System.Numerics
open System.Text.RegularExpressions
open System.Threading.Tasks
open PSWordCloud.Extensions
open PSWordCloud.NewWordCloudCommandHelper
open PSWordCloud.Randomizer
open PSWordCloud.Utils
open SkiaSharp
open SkiaSharp
open SkiaSharp


[<Cmdlet(VerbsCommon.New, "WordCloud", DefaultParameterSetName = "ColorBackground")>]
[<Alias("nwc", "wcloud", "newcloud")>]
type NewWordCloudCommand() =
    inherit PSCmdlet()

    let mutable _resolvedPath = String.Empty
    let mutable _resolvedBackgroundPath = String.Empty

    let mutable _colors : SKColor list = []
    let mutable _fontScale = 1.0f

    let mutable _wordProcessingTasks : Task<string list> list = []
    let _progressId = NextInt()

    member private self.NextColor
        with get() =
            match _colors with
            | head :: tail ->
                _colors <- tail
                head
            | [] ->
                match self.ColorSet with
                | head :: tail ->
                    _colors <- tail
                    head
                | [] -> SKColors.Red

    member private self.NextOrientation =
        if not self.DisableRotation.IsPresent then
            match Randomizer.NextSingle() with
            | x when x > 0.75f -> WordOrientation.Vertical
            | x when x > 0.5f -> WordOrientation.FlippedVertical
            | _ -> WordOrientation.Horizontal
        else
            WordOrientation.Horizontal

    member private self.PaddingMultiplier
        with get() = self.Padding * PaddingBaseScale

    //#region Private Functions

    member private self.ProcessInputAsync (stringLines : string list) =
        seq {
            for line in stringLines do
                yield async {
                    let words =
                        line.Split(SplitChars |> List.toArray, StringSplitOptions.RemoveEmptyEntries)
                        |> Array.toList
                        |> List.choose(
                            fun x ->
                                if (not self.AllowStopWords.IsPresent && StopWords.Contains(x, StringComparer.OrdinalIgnoreCase))
                                    || Regex.Replace(x, "[^a-z-]", String.Empty, RegexOptions.IgnoreCase).Length < 2
                                then
                                    Some x
                                else
                                    None)
                    return words
                }
        } |> Seq.toList

    //#endregion Private Functions

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

    //#region Overrides

    override self.BeginProcessing() =
        if self.MyInvocation.BoundParameters.ContainsKey("RandomSeed") then SetSeed self.RandomSeed

        _colors <- PrepareColorSet self.ColorSet self.BackgroundColor self.StrokeColor self.MaxRenderedWords self.Monochrome.IsPresent
            |> Seq.sortByDescending (fun x -> x.SortValue(NextSingle()))
            |> Seq.toList

    override self.ProcessRecord() =
        let text =
            if self.MyInvocation.ExpectingInput then [ self.InputObject.BaseObject |> To<string> ]
            else
                match self.InputObject.BaseObject with
                | :? (string []) as arr -> arr |> Array.toList
                | :? (Object []) as objArr ->
                    objArr
                    |> Array.toList |> List.map (fun x -> x |> To<string>)
                | x -> [ x.ToString() ]

        _wordProcessingTasks <-
            self.ProcessInputAsync text
            |> List.map Async.StartAsTask
            |> List.append _wordProcessingTasks

    override self.EndProcessing() =
        let allTasks = _wordProcessingTasks |> Task.WhenAll
        let wordScaleDictionary = Dictionary<string, single>(StringComparer.OrdinalIgnoreCase)

        let mutable wordCount = 0
        let mutable background : SKBitmap = null


        allTasks.Wait()
        for lineWords in allTasks.Result do
            CountWords lineWords wordScaleDictionary

        let highestFrequency =
            if self.MyInvocation.BoundParameters.ContainsKey("FocusWord") then
                let maxFreq = wordScaleDictionary.Values.Max() * FocusWordScale
                wordScaleDictionary.[self.FocusWord] <- maxFreq
                maxFreq
            else
                wordScaleDictionary.Values.Max()

        let sortedWords =
            wordScaleDictionary
            |> SortWordList <| self.MaxRenderedWords
            |> Seq.toList

        try
            let cloudBounds =
                if self.MyInvocation.BoundParameters.ContainsKey("BackgroundImage") then
                    background <- SKBitmap.Decode(_resolvedBackgroundPath)
                    SKRectI(0, 0, background.Width, background.Height)
                else
                    SKRectI(0, 0, self.ImageSize.Width, self.ImageSize.Height)
                |> To<SKRect>

            use mutable wordPath = new SKPath()
            use clipRegion = new SKRegion()

            SKRectI.Ceiling(cloudBounds)
            |> clipRegion.SetRect
            |> ignore

            _fontScale <-
                clipRegion.Bounds
                |> To<SKRect>
                |> AdjustFontScale <| wordScaleDictionary.Values.Average()
                                   <| sortedWords.Length

            let scaledWordSizes = Dictionary<string, single>(sortedWords.Length, StringComparer.OrdinalIgnoreCase)
            let maxWordWidth =
                if self.DisableRotation.IsPresent then
                    single cloudBounds.Width * MaxPercentWidth
                else
                    MaxPercentWidth * (Math.Max(cloudBounds.Width, cloudBounds.Height) |> single)

            let mutable retry = false
            let cloudMaxArea = cloudBounds.Width * cloudBounds.Height |> To<single>
            use brush = new SKPaint()
            brush.Typeface <- self.Typeface

            // Pre-test word sizes to scale to image size
            while retry do
                retry <- false
                let size = wordScaleDictionary.[sortedWords.[0]]
                           |> AdjustWordSize <| _fontScale
                                             <| wordScaleDictionary
                brush.DefaultWord size self.StrokeWidth |> ignore

                let mutable wordRect = SKRect.Empty
                brush.MeasureText(sortedWords.[0], ref wordRect) |> ignore
                if wordRect.Width * wordRect.Height * 8.0f < cloudMaxArea * 0.75f then
                    retry <- true
                    _fontScale <- _fontScale * 1.05f

            // Apply user-selected scaling
            _fontScale <- self.WordScale * _fontScale
            retry <- true

            while retry do
                retry <- false
                for word in sortedWords do
                    if not retry then
                        let size = wordScaleDictionary.[sortedWords.[0]]
                                   |> AdjustWordSize <| _fontScale
                                                     <| wordScaleDictionary
                        brush.DefaultWord size self.StrokeWidth |> ignore

                        let mutable wordRect = SKRect.Empty
                        brush.MeasureText(word, ref wordRect) |> ignore

                        if wordRect.Width > maxWordWidth
                            || wordRect.Width * wordRect.Height * 8.0f > cloudMaxArea * 0.75f
                        then
                            retry <- true
                            _fontScale <- _fontScale * 0.98f
                            scaledWordSizes.Clear()

            let aspectRatio = (single cloudBounds.Width) / (single cloudBounds.Height)
            let centre = SKPoint(single cloudBounds.MidX, single cloudBounds.MidY)

            let maxRadius = 0.5f * Math.Max(single cloudBounds.Width, single cloudBounds.Height)

            use writeStream = new SKFileWStream(_resolvedPath)
            use xmlWriter = new SKXmlStreamWriter(writeStream)
            use canvas = SKSvgCanvas.Create(cloudBounds, xmlWriter)
            use filledSpace = new SKRegion()

            if self.AllowOverflow.IsPresent then
                (cloudBounds.Width * BleedAreaScale, cloudBounds.Height * BleedAreaScale)
                |> cloudBounds.Inflate

            if self.ParameterSetName.StartsWith("FileBackground") then
                canvas.DrawBitmap(background, 0.0f, 0.0f)
            elif (self.BackgroundColor <> SKColors.Transparent) then
                canvas.Clear(self.BackgroundColor)

            brush.IsAutohinted <- true
            brush.IsAntialias <- true
            brush.Typeface <- self.Typeface

            let wordProgress = ProgressRecord(_progressId, "Drawing word cloud...", "Finding space for word...")
            let pointProgress = ProgressRecord(_progressId + 1, "Scanning available space...", "Scanning radial points...")
            pointProgress.ParentActivityId <- _progressId

            for word in sortedWords do

            ()
        with
        | e ->
            ErrorRecord(e, "PSWordCloud.GenericError", ErrorCategory.NotSpecified, null)
            |> self.ThrowTerminatingError

    //#endregion Overrides
