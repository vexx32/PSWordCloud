using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Numerics;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using SkiaSharp;

namespace PSWordCloud
{
    /// <summary>
    /// Defines the New-WordCloud cmdlet.
    ///
    /// This command can be used to input large amounts of text, and will generate a word cloud based on
    /// the relative frequencies of the words in the input text.
    /// </summary>
    [Cmdlet(VerbsCommon.New, "WordCloud", DefaultParameterSetName = "ColorBackground")]
    [Alias("wordcloud", "nwc", "wcloud")]
    public class NewWordCloudCommand : PSCmdlet
    {

        #region Constants

        private const float FOCUS_WORD_SCALE = 1.15f;
        private const float BLEED_AREA_SCALE = 1.15f;
        private const float MIN_SATURATION_VALUE = 5f;
        private const float MIN_BRIGHTNESS_DISTANCE = 25f;
        private const float MAX_WORD_WIDTH_PERCENT = 0.75f;
        private const float PADDING_BASE_SCALE = 0.05f;

        internal const float STROKE_BASE_SCALE = 0.02f;

        #endregion Constants

        #region static members

        // TBD

        #endregion static members

        #region Parameters

        /// <summary>
        /// Gets or sets the input text to supply to the word cloud. All input is accepted, but will be treated
        /// as string data regardless of the input type. If you are entering complex object input, ensure they
        /// have a meaningful ToString() method override defined.
        /// </summary>
        /// <value>Accepts piped input or direct array input.</value>
        [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "ColorBackground")]
        [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "ColorBackground-Mono")]
        [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "FileBackground")]
        [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "FileBackground-Mono")]
        [Alias("InputString", "Text", "String", "Words", "Document", "Page")]
        [AllowEmptyString()]
        public PSObject InputObject { get; set; }

        private string _resolvedPath;
        /// <summary>
        /// Gets or sets the output path to save the final SVG vector file to.
        /// </summary>
        /// <value>Accepts a single relative or absolute path as astring.</value>
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "ColorBackground")]
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "ColorBackground-Mono")]
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "FileBackground")]
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "FileBackground-Mono")]
        [Alias("OutFile", "ExportPath", "ImagePath")]
        public string Path
        {
            get => _resolvedPath;
            set => _resolvedPath = SessionState.Path
                .GetUnresolvedProviderPathFromPSPath(value, out ProviderInfo provider, out PSDriveInfo drive);
        }

        private string _backgroundFullPath;
        /// <summary>
        /// Gets or sets the path to the background image to be used as a base for the final word cloud image.
        /// </summary>
        /// <value>Accepts a single relative or absolute path as astring.</value>
        [Parameter(Mandatory = true, ParameterSetName = "FileBackground")]
        [Parameter(Mandatory = true, ParameterSetName = "FileBackground-Mono")]
        public string BackgroundImage
        {
            get => _backgroundFullPath;
            set
            {
                var previousDir = Environment.CurrentDirectory;
                Environment.CurrentDirectory = SessionState.Path.CurrentFileSystemLocation.Path;
                _backgroundFullPath = System.IO.Path.GetFullPath(value);
            }

        }

        /// <summary>
        /// Gets or sets the image size value for the word cloud image.
        ///
        /// Input can be passed either directly as a SkiaSharp.SKSizeI object, or in one of the following
        /// string alternate forms:
        ///
        ///     - A predefined size string. One of:
        ///         * 720p           (canvas size: 1280x720)
        ///         * 1080p          (canvas size: 1920x1080)
        ///         * 4K             (canvas size: 3840x2160)
        ///         * A4             (canvas size: 816x1056)
        ///         * Poster11x17    (canvas size: 1056x1632)
        ///         * Poster18x24    (canvas size: 1728x2304)
        ///         * Poster24x36    (canvas size: 2304x3456)
        ///     - Single integer (e.g., -ImageSize 1024). This will be used as both the width and height of the
        ///       image, creating a square canvas.
        ///     - Any image size string (e.g., 1024x768). The first number will be used as the width, and the
        ///       second number used as the height of the canvas.
        ///     - A hashtable or custom object with keys or properties named "Width" and "Height" that contain
        ///       integer values
        /// </summary>
        /// <value></value>
        [Parameter(ParameterSetName = "ColorBackground")]
        [Parameter(ParameterSetName = "ColorBackground-Mono")]
        [ArgumentCompleter(typeof(ImageSizeCompleter))]
        [TransformToSKSizeI]
        public SKSizeI ImageSize { get; set; } = new SKSizeI(4096, 2304);

        [Parameter]
        [Alias("FontFamily", "FontFace")]
        [ArgumentCompleter(typeof(FontFamilyCompleter))]
        [TransformToSKTypeface]
        public SKTypeface Typeface { get; set; } = WCUtils.FontManager.MatchFamily(
            "Consolas", SKFontStyle.Normal);

        [Parameter(ParameterSetName = "ColorBackground")]
        [Parameter(ParameterSetName = "ColorBackground-Mono")]
        [Alias("Backdrop", "CanvasColor")]
        [ArgumentCompleter(typeof(SKColorCompleter))]
        [TransformToSKColor]
        public SKColor BackgroundColor { get; set; } = SKColors.Black;

        private List<SKColor> _colors;
        [Parameter]
        [TransformToSKColor]
        [ArgumentCompleter(typeof(SKColorCompleter))]
        public SKColor[] ColorSet { get; set; } = WCUtils.StandardColors.ToArray();

        [Parameter]
        [Alias("OutlineWidth")]
        [ValidateRange(0, 10)]
        public int StrokeWidth { get; set; } = 0;

        [Parameter]
        [Alias("OutlineColor")]
        [TransformToSKColor]
        [ArgumentCompleter(typeof(SKColorCompleter))]
        public SKColor StrokeColor { get; set; } = SKColors.Black;

        [Parameter]
        [Alias("Title")]
        public string FocusWord { get; set; }

        [Parameter]
        [Alias("ScaleFactor")]
        [ValidateRange(0.01, 20)]
        public float WordScale { get; set; } = 1;

        [Parameter]
        [Alias("Spacing")]
        public float Padding { get; set; } = 5;

        [Parameter]
        [ValidateRange(1, 500)]
        public float DistanceStep { get; set; } = 5;

        [Parameter]
        [ValidateRange(1, 50)]
        public float RadialStep { get; set; } = 15;

        [Parameter]
        [Alias("MaxWords")]
        [ValidateRange(0, int.MaxValue)]
        public int MaxRenderedWords { get; set; } = 100;

        [Parameter]
        [Alias("MaxColours")]
        [ValidateRange(1, int.MaxValue)]
        public int MaxColors { get; set; } = int.MaxValue;

        [Parameter]
        [Alias("SeedValue")]
        public int RandomSeed { get; set; }

        [Parameter]
        [Alias("DisableWordRotation")]
        public SwitchParameter DisableRotation { get; set; }

        [Parameter(Mandatory = true, ParameterSetName = "FileBackground-Mono")]
        [Parameter(Mandatory = true, ParameterSetName = "ColorBackground-Mono")]
        [Alias("BlackAndWhite", "Greyscale")]
        public SwitchParameter Monochrome { get; set; }

        [Parameter]
        [Alias("IgnoreStopWords")]
        public SwitchParameter AllowStopWords { get; set; }

        [Parameter]
        public SwitchParameter PassThru { get; set; }

        [Parameter()]
        [Alias("AllowBleed")]
        public SwitchParameter AllowOverflow { get; set; }

        #endregion Parameters

        #region staticMembers

        private static readonly string[] _stopWords = new[] {
            "a","about","above","after","again","against","all","am","an","and","any","are","aren't","as","at","be",
            "because","been","before","being","below","between","both","but","by","can't","cannot","could","couldn't",
            "did","didn't","do","does","doesn't","doing","don't","down","during","each","few","for","from","further",
            "had","hadn't","has","hasn't","have","haven't","having","he","he'd","he'll","he's","her","here","here's",
            "hers","herself","him","himself","his","how","how's","i","i'd","i'll","i'm","i've","if","in","into","is",
            "isn't","it","it's","its","itself","let's","me","more","most","mustn't","my","myself","no","nor","not","of",
            "off","on","once","only","or","other","ought","our","ours","ourselves","out","over","own","same","shan't",
            "she","she'd","she'll","she's","should","shouldn't","so","some","such","than","that","that's","the","their",
            "theirs","them","themselves","then","there","there's","these","they","they'd","they'll","they're","they've",
            "this","those","through","to","too","under","until","up","very","was","wasn't","we","we'd","we'll","we're",
            "we've","were","weren't","what","what's","when","when's","where","where's","which","while","who","who's",
            "whom","why","why's","with","won't","would","wouldn't","you","you'd","you'll","you're","you've","your",
            "yours","yourself","yourselves" };

        private static readonly char[] _splitChars = new[] {
            ' ','\n','\t','\r','.',',',';','\\','/','|',
            ':','"','?','!','{','}','[',']',':','(',')',
            '<','>','“','”','*','#','%','^','&','+','=' };

        private static Random _random;
        private static Random Random { get => _random = _random ?? new Random(); }
        private static float RandomFloat { get => (float)Random.NextDouble(); }

        #endregion staticMembers

        #region privateVariables

        private List<Task<IEnumerable<string>>> _wordProcessingTasks;
        private float _fontScale;
        private int _progressID;
        private int _colorIndex = 0;
        private SKColor _nextColor
        {
            get
            {
                if (_colorIndex == _colors.Count())
                {
                    _colorIndex = 0;
                }

                var color = _colors[_colorIndex];
                _colorIndex++;

                return color;
            }
        }
        private float _paddingMultiplier
        {
            get => Padding * PADDING_BASE_SCALE;
        }

        #endregion privateVariables

        protected override void BeginProcessing()
        {
            _random = MyInvocation.BoundParameters.ContainsKey(nameof(RandomSeed))
                ? new Random(RandomSeed)
                : new Random();
            _progressID = Random.Next();

            _colors = ProcessColorSet(ColorSet, BackgroundColor, StrokeColor, MaxRenderedWords, Monochrome)
                .OrderByDescending(x => x.SortValue(RandomFloat))
                .ToList();
        }


        protected override void ProcessRecord()
        {
            string[] text;
            if (MyInvocation.ExpectingInput)
            {
                text = new[] { InputObject.BaseObject as string };
            }
            else
            {
                text = InputObject.BaseObject as string[] ?? new[] { InputObject.BaseObject as string };
            }

            if (_wordProcessingTasks == null)
            {
                _wordProcessingTasks = new List<Task<IEnumerable<string>>>(text.Length);
            }

            foreach (var line in text)
            {
                _wordProcessingTasks.Add(ProcessInputAsync(line));
            }
        }

        protected override void EndProcessing()
        {
            var lineStrings = Task.WhenAll<IEnumerable<string>>(_wordProcessingTasks);
            lineStrings.Wait();

            var wordCount = 0;
            var wordScaleDictionary = new Dictionary<string, float>(StringComparer.OrdinalIgnoreCase);
            SKRect wordBounds = SKRect.Empty;
            SKRect drawableBounds;
            SKPath wordPath = null;
            SKRegion clipRegion = null;
            SKBitmap backgroundImage = null;
            ProgressRecord wordProgress = null;
            ProgressRecord pointProgress = null;

            float inflationValue = 0;

            foreach (var lineWords in lineStrings.Result)
            {
                CountWords(lineWords, wordScaleDictionary);
            }

            // All words counted and in the dictionary.
            var highestWordFreq = wordScaleDictionary.Values.Max();

            if (MyInvocation.BoundParameters.ContainsKey(nameof(FocusWord)))
            {
                wordScaleDictionary[FocusWord] = highestWordFreq = highestWordFreq * FOCUS_WORD_SCALE;
            }

            List<string> sortedWordList = new List<string>(SortWordList(wordScaleDictionary, MaxRenderedWords));

            try
            {
                if (MyInvocation.BoundParameters.ContainsKey(nameof(BackgroundImage)))
                {
                    backgroundImage = SKBitmap.Decode(_backgroundFullPath);
                    drawableBounds = new SKRectI(0, 0, backgroundImage.Width, backgroundImage.Height);
                }
                else
                {
                    drawableBounds = new SKRectI(0, 0, ImageSize.Width, ImageSize.Height);
                }

                wordPath = new SKPath();
                clipRegion = new SKRegion();
                clipRegion.SetRect(SKRectI.Round(drawableBounds));

                _fontScale = FontScale(
                    clipRegion.Bounds,
                    WordScale,
                    wordScaleDictionary.Values.Average(),
                    Math.Min(wordScaleDictionary.Count, MaxRenderedWords));

                var scaledWordSizes = new Dictionary<string, float>(
                    sortedWordList.Count, StringComparer.OrdinalIgnoreCase);
                float maxWordWidth = DisableRotation
                    ? drawableBounds.Width * MAX_WORD_WIDTH_PERCENT
                    : Math.Max(drawableBounds.Width, drawableBounds.Height) * MAX_WORD_WIDTH_PERCENT;

                bool retry;
                using (SKPaint brush = new SKPaint())
                {
                    brush.Typeface = Typeface;
                    SKRect rect = SKRect.Empty;
                    float adjustedWordSize;

                    do
                    {
                        // Pre-test and adjust global scale based on the largest word.
                        retry = false;
                        adjustedWordSize = ScaleWordSize(
                            wordScaleDictionary[sortedWordList[0]], _fontScale, wordScaleDictionary);
                        brush.NextWord(adjustedWordSize, StrokeWidth);
                        var adjustedTextWidth = brush.MeasureText(sortedWordList[0], ref rect) * (1 + _paddingMultiplier);
                        if ((rect.Width * rect.Height * 8) < (drawableBounds.Width * drawableBounds.Height * 0.75f))
                        {
                            retry = true;
                            _fontScale *= 1.05f;
                        }
                    } while (retry);

                    // Apply manual scaling from the user
                    _fontScale *= WordScale;

                    do
                    {
                        retry = false;
                        foreach (string word in sortedWordList)
                        {
                            adjustedWordSize = ScaleWordSize(
                                wordScaleDictionary[word], _fontScale, wordScaleDictionary);

                            brush.NextWord(adjustedWordSize, StrokeWidth);
                            var adjustedTextWidth = brush.MeasureText(word, ref rect)
                                * (1 + _paddingMultiplier + StrokeWidth * STROKE_BASE_SCALE);

                            if (adjustedTextWidth > maxWordWidth
                                || rect.Width * rect.Height * 8 > drawableBounds.Width * drawableBounds.Height * 0.75f)
                            {
                                retry = true;
                                _fontScale *= 0.95f;
                                scaledWordSizes.Clear();
                                break;
                            }

                            scaledWordSizes[word] = adjustedWordSize;
                        }
                    }
                    while (retry);
                }

                var aspectRatio = drawableBounds.Width / (float)drawableBounds.Height;
                SKPoint centrePoint = new SKPoint(drawableBounds.MidX, drawableBounds.MidY);

                // Remove all words that were cut from the final rendering list
                sortedWordList.RemoveAll(x => !scaledWordSizes.ContainsKey(x));

                var maxRadius = Math.Max(drawableBounds.Width, drawableBounds.Height) / 2f;

                using (SKFileWStream outputStream = new SKFileWStream(_resolvedPath))
                using (SKXmlStreamWriter xmlWriter = new SKXmlStreamWriter(outputStream))
                using (SKCanvas canvas = SKSvgCanvas.Create(drawableBounds, xmlWriter))
                using (SKPaint brush = new SKPaint())
                using (SKRegion occupiedSpace = new SKRegion())
                {
                    if (ParameterSetName.StartsWith("FileBackground"))
                    {
                        canvas.DrawBitmap(backgroundImage, 0, 0);
                    }
                    else if (BackgroundColor != SKColor.Empty)
                    {
                        canvas.Clear(BackgroundColor);
                    }

                    WordOrientation targetOrientation;
                    SKPoint targetPoint;

                    brush.IsAutohinted = true;
                    brush.IsAntialias = true;
                    brush.Typeface = Typeface;

                    wordProgress = new ProgressRecord(
                        _progressID,
                        "Drawing word cloud...",
                        "Finding space for word...");
                    pointProgress = new ProgressRecord(
                        _progressID + 1,
                        "Scanning available space...",
                        "Scanning radial points...");
                    pointProgress.ParentActivityId = _progressID;

                    foreach (string word in sortedWordList.OrderByDescending(x => scaledWordSizes[x]))
                    {
                        wordCount++;

                        inflationValue = scaledWordSizes[word] * (_paddingMultiplier + StrokeWidth * STROKE_BASE_SCALE);
                        targetOrientation = WordOrientation.Horizontal;
                        targetPoint = SKPoint.Empty;

                        var wordColor = _nextColor;
                        brush.NextWord(scaledWordSizes[word], StrokeWidth, wordColor);

                        wordPath.Reset();
                        wordBounds = SKRect.Empty;
                        brush.MeasureText(word, ref wordBounds);

                        var wordWidth = wordBounds.Width;
                        var wordHeight = wordBounds.Height;

                        var percentComplete = 100f * wordCount / scaledWordSizes.Count;

                        wordProgress.StatusDescription = string.Format(
                            "Draw: \"{0}\" [Size: {1:0}] ({2} of {3})",
                            word, brush.TextSize, wordCount, scaledWordSizes.Count);
                        wordProgress.PercentComplete = (int)Math.Round(percentComplete);
                        WriteProgress(wordProgress);

                        for (
                            float radius = 0;
                            radius <= maxRadius;
                            radius += GetRadiusIncrement(
                                scaledWordSizes[word], DistanceStep, maxRadius, inflationValue, percentComplete))
                        {
                            SKPoint adjustedPoint, baseOffset;

                            var radialPoints = GetRadialPoints(centrePoint, radius, RadialStep, aspectRatio);
                            var totalPoints = radialPoints.Count();
                            var pointsChecked = 0;
                            foreach (var point in radialPoints)
                            {
                                pointsChecked++;
                                if (!drawableBounds.Contains(point) && point != centrePoint)
                                {
                                    continue;
                                }

                                var orientation = _wordOrientation;
                                pointProgress.Activity = string.Format(
                                    "Finding available space to draw with orientation: {0}",
                                    orientation);
                                pointProgress.StatusDescription = string.Format(
                                    "Checking [Point:{0,8:N2}, {1,8:N2}] ({2,4} / {3,4}) at [Radius: {4,8:N2}]",
                                    point.X, point.Y, pointsChecked, totalPoints, radius);
                                //pointProgress.PercentComplete = 100 * pointsChecked / totalPoints;
                                WriteProgress(pointProgress);

                                baseOffset = new SKPoint(
                                    -(wordWidth / 2),
                                    (wordHeight / 2));
                                adjustedPoint = point + baseOffset;

                                SKMatrix rotation;
                                rotation = GetRotationMatrix(point, orientation);

                                SKPath alteredPath = brush.GetTextPath(word, adjustedPoint.X, adjustedPoint.Y);
                                alteredPath.Transform(rotation);
                                alteredPath.GetTightBounds(out wordBounds);

                                wordBounds.Inflate(inflationValue * 2, inflationValue * 2);

                                if (wordBounds.FallsOutside(clipRegion))
                                {
                                    continue;
                                }

                                if (word == FocusWord || !occupiedSpace.IntersectsRect(wordBounds))
                                {
                                    wordPath = alteredPath;
                                    targetPoint = adjustedPoint;
                                    targetOrientation = orientation;
                                    goto nextWord;
                                }

                                if (radius == 0)
                                {
                                    // No point checking more than a single point at the origin
                                    break;
                                }
                            }
                        }

                    nextWord:
                        if (targetPoint != SKPoint.Empty)
                        {
                            wordPath.FillType = SKPathFillType.EvenOdd;
                            if (MyInvocation.BoundParameters.ContainsKey(nameof(StrokeWidth)))
                            {
                                brush.Color = StrokeColor;
                                brush.IsStroke = true;
                                brush.Style = SKPaintStyle.Stroke;
                                canvas.DrawPath(wordPath, brush);

                            }

                            brush.IsStroke = false;
                            brush.Color = wordColor;
                            brush.Style = SKPaintStyle.Fill;
                            occupiedSpace.Op(wordPath, SKRegionOperation.Union);
                            canvas.DrawPath(wordPath, brush);
                        }
                    }

                    canvas.Flush();
                    outputStream.Flush();
                    if (PassThru.IsPresent)
                    {
                        WriteObject(InvokeProvider.Item.Get(_resolvedPath), true);
                    }
                }
            }
            catch (Exception e)
            {
                WriteError(new ErrorRecord(e, "PSWordCloudError", ErrorCategory.InvalidResult, null));
            }
            finally
            {
                clipRegion?.Dispose();
                wordPath?.Dispose();
                backgroundImage?.Dispose();

                if (wordProgress != null)
                {
                    wordProgress.RecordType = ProgressRecordType.Completed;
                    WriteProgress(wordProgress);
                }

                if (pointProgress != null)
                {
                    pointProgress.RecordType = ProgressRecordType.Completed;
                    WriteProgress(pointProgress);
                }
            }
        }

        private static SKMatrix GetRotationMatrix(SKPoint point, WordOrientation orientation)
        {
            switch (orientation)
            {
                case WordOrientation.Vertical:
                    return SKMatrix.MakeRotationDegrees(90, point.X, point.Y);

                case WordOrientation.FlippedVertical:
                    return SKMatrix.MakeRotationDegrees(270, point.X, point.Y);

                default:
                    return SKMatrix.MakeIdentity();
            }
        }

        private static IEnumerable<SKColor> ProcessColorSet(
            SKColor[] set, SKColor background, SKColor stroke, int maxCount, bool monochrome)
        {
            Random.Shuffle(set);
            background.ToHsv(out float bh, out float bs, out float backgroundBrightness);

            foreach (var color in set.Where(x => x != stroke && x != background).Take(maxCount))
            {
                if (!monochrome)
                {
                    color.ToHsv(out float h, out float s, out float v);
                    if (s >= MIN_SATURATION_VALUE && Math.Abs(v - backgroundBrightness) > MIN_BRIGHTNESS_DISTANCE)
                    {
                        yield return color;
                    }
                }
                else
                {
                    color.ToHsv(out float h, out float s, out float brightness);
                    byte level = (byte)Math.Floor(255 * brightness / 100f);
                    yield return new SKColor(level, level, level);
                }
            }
        }

        private static void CountWords(IEnumerable<string> wordList, IDictionary<string, float> dictionary)
        {
            foreach (string word in wordList)
            {
                var trimmedWord = Regex.Replace(word, "s$", string.Empty, RegexOptions.IgnoreCase);
                var pluralWord = string.Format("{0}s", word);
                if (dictionary.ContainsKey(trimmedWord))
                {
                    dictionary[trimmedWord]++;
                }
                else if (dictionary.ContainsKey(pluralWord))
                {
                    dictionary[word] = dictionary[pluralWord] + 1;
                    dictionary.Remove(pluralWord);
                }
                else
                {
                    dictionary[word] = dictionary.ContainsKey(word) ? dictionary[word] + 1 : 1;
                }
            }
        }

        private WordOrientation _wordOrientation
        {
            get
            {
                if (!DisableRotation)
                {
                    var num = RandomFloat;
                    if (num > 0.75)
                    {
                        return WordOrientation.Vertical;
                    }

                    if (num > 0.5)
                    {
                        return WordOrientation.FlippedVertical;
                    }
                }

                return WordOrientation.Horizontal;
            }
        }

        private static float FontScale(SKRect space, float baseScale, float averageWordFrequency, int wordCount)
        {
            return baseScale * (space.Height + space.Width)
                / (8 * averageWordFrequency * wordCount);
        }

        private static float ScaleWordSize(
            float baseSize, float globalScale, IDictionary<string, float> scaleDictionary)
        {
            return baseSize * globalScale * (2 * RandomFloat
                / (1 + scaleDictionary.Values.Max() - scaleDictionary.Values.Min()) + 0.9f);
        }

        private static IEnumerable<string> SortWordList(IDictionary<string, float> dictionary, int maxWords)
        {
            return dictionary.Keys.OrderByDescending(word => dictionary[word])
                .Take(maxWords == 0 ? int.MaxValue : maxWords);
        }

        private static float GetRadiusIncrement(
            float wordSize, float distanceStep, float maxRadius, float padding, float percentComplete)
            => (5 + RandomFloat * (2.5f + percentComplete / 10)) * distanceStep * wordSize * (1 + padding) / maxRadius;

        private static IEnumerable<SKPoint> GetRadialPoints(
            SKPoint centre, float radius, float radialStep, float aspectRatio = 1)
        {
            if (radius == 0)
            {
                yield return centre;
                yield break;
            }

            Complex point;
            float angle = 0, maxAngle = 0, angleIncrement = 360 / (radius / 6 + 1);
            bool clockwise = RandomFloat > 0.5;

            switch (Random.Next() % 4)
            {
                case 0:
                    angle = 0;
                    break;

                case 1:
                    angle = 90;
                    break;

                case 2:
                    angle = 180;
                    break;

                case 3:
                    angle = 270;
                    break;
            }

            if (clockwise)
            {
                maxAngle = angle + 360;
            }
            else
            {
                maxAngle = angle - 360;
                angleIncrement *= -1;
            }

            do
            {
                point = Complex.FromPolarCoordinates(radius, angle.ToRadians());
                yield return new SKPoint(centre.X + (float)point.Real * aspectRatio, centre.Y + (float)point.Imaginary);

                angle += angleIncrement;
            } while (clockwise ? angle <= maxAngle : angle >= maxAngle);
        }

        private async Task<IEnumerable<string>> ProcessInputAsync(string line)
        {
            return await Task.Run<IEnumerable<string>>(
                () =>
                {
                    var words = new List<string>(line.Split(_splitChars, StringSplitOptions.RemoveEmptyEntries));
                    words.RemoveAll(
                        x => (!AllowStopWords && _stopWords.Contains(x, StringComparer.OrdinalIgnoreCase))
                            || Regex.Replace(x, "[^a-z-]", string.Empty, RegexOptions.IgnoreCase).Length < 2);
                    return words;
                });
        }
    }
}
