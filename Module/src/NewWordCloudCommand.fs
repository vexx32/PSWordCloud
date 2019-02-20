namespace PSWordCloud
open System
open System.Collections.Generic
open System.Linq
open System.Management.Automation
open System.Text.RegularExpressions
open System.Threading.Tasks
open PSWordCloud.Extensions
open PSWordCloud.NewWordCloudCommandHelper
open PSWordCloud.Randomizer
open PSWordCloud.Utils
open SkiaSharp


[<Cmdlet(VerbsCommon.New, "WordCloud", DefaultParameterSetName = "ColorBackground")>]
[<Alias("nwc", "wcloud", "newcloud")>]
type NewWordCloudCommand() =
    inherit PSCmdlet()

    let mutable _resolvedPath = String.Empty
    let mutable _resolvedBackgroundPath = String.Empty

    let mutable _colors : SKColor list = []

    let mutable _wordProcessingTasks : Task<string list> list = []
    let _progressId = NextInt()

    member private self.NextColor
        with get() =
            match _colors with
            | head :: tail ->
                _colors <- tail
                head
            | [] ->
                match self.ColorSet |> Array.toList with
                | head :: tail ->
                    _colors <- tail
                    head
                | [] -> SKColors.Red

    member private self.NextOrientation
        with get() =
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

    member private self.ProcessInputAsync stringLines =
        seq {
            for (line : string) in stringLines do
                yield async {
                    let words =
                        line.Split(SplitChars, StringSplitOptions.RemoveEmptyEntries)
                        |> Array.toList
                        |> List.choose (fun x ->
                            if
                                (not self.AllowStopWords.IsPresent && StopWords.Contains(x, StringComparer.OrdinalIgnoreCase))
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
    member val public ColorSet = StandardColors |> Seq.toArray
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

        let wordCount = ref 0
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
            |> SortWordList self.MaxRenderedWords
            |> Seq.toList

        let wordProgress = ProgressRecord(_progressId, "Drawing word cloud...", "Finding space for word...")
        let pointProgress = ProgressRecord(_progressId + 1, "Scanning available space...", "Scanning radial points...")
        pointProgress.ParentActivityId <- _progressId

        try
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

                FontScale <-
                    clipRegion.Bounds
                    |> To<SKRect>
                    |> AdjustFontScale <| wordScaleDictionary.Values.Average()
                                       <| sortedWords.Length

                let maxWordWidth =
                    if self.DisableRotation.IsPresent then
                        single cloudBounds.Width * MaxPercentWidth
                    else
                        MaxPercentWidth * (Math.Max(cloudBounds.Width, cloudBounds.Height) |> single)

                let cloudMaxArea =
                    if not self.AllowOverflow.IsPresent then
                        cloudBounds.Width * cloudBounds.Height
                    else
                        cloudBounds.Width * cloudBounds.Height * 1.5f
                let aspectRatio = (single cloudBounds.Width) / (single cloudBounds.Height)

                // Pre-test word sizes to scale to image size
                setBaseFontScale wordScaleDictionary sortedWords.[0] self.StrokeWidth cloudMaxArea

                // Apply user-selected scaling
                FontScale <- self.WordScale * FontScale

                let wordSizes = getWordScaleDictionary wordScaleDictionary sortedWords maxWordWidth aspectRatio self.StrokeWidth self.AllowOverflow.IsPresent

                let centre = SKPoint(single cloudBounds.MidX, single cloudBounds.MidY)

                let maxRadius = 0.5f * Math.Max(single cloudBounds.Width, single cloudBounds.Height)

                use writeStream = new SKFileWStream(_resolvedPath)
                use xmlWriter = new SKXmlStreamWriter(writeStream)
                use canvas = SKSvgCanvas.Create(cloudBounds, xmlWriter)
                use filledSpace = new SKRegion()
                use brush = new SKPaint()

                if self.AllowOverflow.IsPresent then
                    cloudBounds.Inflate(cloudBounds.Width * BleedAreaScale, cloudBounds.Height * BleedAreaScale)

                if self.ParameterSetName.StartsWith("FileBackground") then
                    canvas.DrawBitmap(background, 0.0f, 0.0f)
                elif (self.BackgroundColor <> SKColors.Transparent) then
                    canvas.Clear(self.BackgroundColor)

                brush.IsAutohinted <- true
                brush.IsAntialias <- true
                brush.Typeface <- self.Typeface

                let statusString = "Draw: \"{0}\" [Size: {1:0}] ({2} of {3})"
                let pointActivity = "Finding available space to draw with orientation: {0}"
                let pointStatus = "Checking [Point:{0,8:N2}, {1,8:N2}] ({2,4} / {3,4}) at [Radius: {4,8:N2}]"

                let totalWords = wordSizes.Count

                for word in sortedWords do
                    incr wordCount
                    let inflationValue = wordSizes.[word] * (self.PaddingMultiplier + self.StrokeWidth * StrokeBaseScale)
                    let wordColor = self.NextColor
                    brush.NextWord wordColor wordSizes.[word] self.StrokeWidth |> ignore

                    wordPath <- brush.GetTextPath(word, 0.0f, 0.0f)
                    let wordBounds = ref <| wordPath.ComputeTightBounds()

                    let wordWidth = wordBounds.Value.Width
                    let wordHeight = wordBounds.Value.Height

                    let mutable targetPoint : SKPoint option = None
                    let mutable orientation = WordOrientation.Horizontal

                    let percentComplete = 100.0f * (single wordCount.Value) / (single totalWords)
                    wordProgress.StatusDescription <- String.Format(statusString, word, brush.TextSize, wordCount.Value, totalWords)
                    wordProgress.PercentComplete <- int <| Math.Round(double percentComplete)
                    self.WriteProgress(wordProgress)

                    let increment = GetRadiusIncrement wordSizes.[word] self.DistanceStep maxRadius inflationValue percentComplete

                    for radius in 0.0f..increment..maxRadius do
                        let radialPoints = GetRadialPoints centre radius self.RadialStep aspectRatio
                        let totalPoints = radialPoints.Count()
                        let pointCount = ref 0

                        if targetPoint = None then
                            for point in radialPoints do
                                incr pointCount
                                if
                                    cloudBounds.Contains(point) || wordCount.Value = 1
                                    && targetPoint = None
                                then
                                    orientation <- self.NextOrientation

                                    pointProgress.Activity <- String.Format(pointActivity, orientation)
                                    pointProgress.StatusDescription <- String.Format(pointStatus, point.X, point.Y, pointCount.Value, totalPoints, radius)
                                    self.WriteProgress(pointProgress)

                                    let baseOffset = SKPoint(-wordWidth / 2.0f, wordHeight / 2.0f)
                                    let adjustedPoint = point + baseOffset

                                    let rotation = orientation |> ToRotationMatrix point
                                    let alteredPath = brush.GetTextPath(word, adjustedPoint.X, adjustedPoint.Y)
                                    alteredPath.Transform(rotation)
                                    alteredPath.GetTightBounds(wordBounds) |> ignore

                                    wordBounds.Value.Inflate(inflationValue, inflationValue)

                                    if
                                        not (wordBounds.Value.FallsOutside clipRegion)
                                        && not (filledSpace.Intersects wordBounds.Value)
                                    then
                                        wordPath <- alteredPath
                                        targetPoint <- Some adjustedPoint

                    if targetPoint.IsSome then
                        wordPath.FillType <- SKPathFillType.EvenOdd

                        if self.MyInvocation.BoundParameters.ContainsKey("StrokeWidth") then
                            brush.Color <- self.StrokeColor
                            brush.IsStroke <- true
                            brush.Style <- SKPaintStyle.Stroke
                            canvas.DrawPath(wordPath, brush)

                        brush.IsStroke <- false
                        brush.Color <- wordColor
                        brush.Style <- SKPaintStyle.Fill
                        canvas.DrawPath(wordPath, brush)
                        filledSpace.Op(wordPath, SKRegionOperation.Union) |> ignore

                canvas.Flush()
                writeStream.Flush()

                if self.PassThru.IsPresent then
                    let item = self.InvokeProvider.Item.Get(_resolvedPath)
                    self.WriteObject(item, true)
            with
            | e ->
                ErrorRecord(e, "PSWordCloud.GenericError", ErrorCategory.NotSpecified, null)
                |> self.ThrowTerminatingError
        finally
            wordProgress.RecordType <- ProgressRecordType.Completed
            wordProgress.RecordType <- ProgressRecordType.Completed
            self.WriteProgress(wordProgress)
            self.WriteProgress(pointProgress)

    //#endregion Overrides
