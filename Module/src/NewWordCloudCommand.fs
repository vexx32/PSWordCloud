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
    let mutable _fontScale = 1.0f

    let mutable _wordProcessingTasks : Task<string list> list = []
    let _progressId = NextInt()

    let rec setBaseFontScale (dictionary : Dictionary<string, single>) largestWord strokeWidth maxArea =
        use brush = new SKPaint()
        let size = dictionary.[largestWord]
                   |> AdjustWordSize <| _fontScale
                                     <| dictionary
        brush.DefaultWord size strokeWidth |> ignore

        let mutable wordRect = SKRect.Empty
        brush.MeasureText(largestWord, ref wordRect) |> ignore
        if wordRect.Width * wordRect.Height * 8.0f < maxArea * 0.75f then
            _fontScale <- _fontScale * 1.05f
            setBaseFontScale dictionary largestWord strokeWidth maxArea

    let rec scaleWords (wordScales : Dictionary<string,single>) (wordSizes : Dictionary<string,single>) (wordList : string list) maxWidth aspect strokeWidth overflow =
        use brush = new SKPaint()
        let maxArea = maxWidth * maxWidth * (if aspect > 1.0f then 1.0f / aspect else aspect)

        match wordList with
        | [] -> ()
        | head :: tail ->
            let size = wordScales.[head]
                       |> AdjustWordSize <| _fontScale
                                         <| wordScales
            brush.DefaultWord size strokeWidth |> ignore

            let mutable wordRect = SKRect.Empty
            brush.MeasureText(head, ref wordRect) |> ignore

            if (wordRect.Width > maxWidth
                || wordRect.Width * wordRect.Height * 8.0f > maxArea * 0.75f)
                && not overflow
            then
                _fontScale <- _fontScale * 0.98f
                wordSizes.Clear()
                scaleWords wordScales wordSizes wordList maxWidth aspect strokeWidth overflow
            else
                wordSizes.[head] <- size
                scaleWords wordScales wordSizes tail maxWidth aspect strokeWidth overflow


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

                let cloudMaxArea =
                    if not self.AllowOverflow.IsPresent then
                        cloudBounds.Width * cloudBounds.Height
                    else
                        cloudBounds.Width * cloudBounds.Height * 1.5f
                let aspectRatio = (single cloudBounds.Width) / (single cloudBounds.Height)
                use brush = new SKPaint()
                brush.Typeface <- self.Typeface

                // Pre-test word sizes to scale to image size
                setBaseFontScale wordScaleDictionary sortedWords.[0] self.StrokeWidth cloudMaxArea

                // Apply user-selected scaling
                _fontScale <- self.WordScale * _fontScale

                scaleWords wordScaleDictionary scaledWordSizes sortedWords maxWordWidth aspectRatio self.StrokeWidth self.AllowOverflow.IsPresent

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

                let statusString = "Draw: \"{0}\" [Size: {1:0}] ({2} of {3})"
                let pointActivity = "Finding available space to draw with orientation: {0}"
                let pointStatus = "Checking [Point:{0,8:N2}, {1,8:N2}] ({2,4} / {3,4}) at [Radius: {4,8:N2}]"

                let totalWords = scaledWordSizes.Count

                for word in sortedWords do
                    incr <| ref wordCount
                    let inflationValue = scaledWordSizes.[word] * (self.PaddingMultiplier + self.StrokeWidth * StrokeBaseScale)
                    let wordColor = self.NextColor
                    brush.NextWord wordColor scaledWordSizes.[word] self.StrokeWidth |> ignore

                    wordPath <- brush.GetTextPath(word, 0.0f, 0.0f)
                    let wordBounds = wordPath.ComputeTightBounds()

                    let wordWidth = wordBounds.Width
                    let wordHeight = wordBounds.Height

                    let mutable targetPoint : SKPoint option = None
                    let mutable orientation = WordOrientation.Horizontal

                    let percentComplete = 100.0f * (single wordCount) / (single totalWords)
                    wordProgress.StatusDescription <-
                        (statusString, word, brush.TextSize, wordCount, totalWords)
                        |> String.Format
                    wordProgress.PercentComplete <- int <| Math.Round(double percentComplete)
                    self.WriteProgress(wordProgress)

                    let increment = GetRadiusIncrement scaledWordSizes.[word] self.DistanceStep maxRadius inflationValue percentComplete

                    for radius in 0.0f..increment..maxRadius do
                        let radialPoints = GetRadialPoints centre radius self.RadialStep aspectRatio
                        let totalPoints = radialPoints.Count()
                        let mutable pointCount = 0

                        if targetPoint = None then
                            for point in radialPoints do
                                incr <| ref pointCount
                                if
                                    cloudBounds.Contains(point) || wordCount = 1
                                    && targetPoint = None
                                then
                                    orientation <- self.NextOrientation

                                    pointProgress.Activity <-
                                        String.Format(pointActivity, orientation)
                                    pointProgress.StatusDescription <-
                                        String.Format(pointStatus, point.X, point.Y, pointCount, totalPoints, radius)
                                    self.WriteProgress(pointProgress)

                                    let baseOffset = SKPoint(-wordWidth / 2.0f, wordHeight / 2.0f)
                                    let adjustedPoint = point + baseOffset

                                    let rotation = orientation |> ToRotationMatrix point
                                    let alteredPath = brush.GetTextPath(word, adjustedPoint.X, adjustedPoint.Y)
                                    alteredPath.Transform(rotation)
                                    alteredPath.GetTightBounds(ref wordBounds) |> ignore

                                    wordBounds.Inflate(inflationValue, inflationValue)

                                    if
                                        not (wordBounds.FallsOutside clipRegion)
                                        && not (filledSpace.Intersects wordBounds)
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
