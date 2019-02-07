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
    [Cmdlet(VerbsCommon.New, "WordCloud", DefaultParameterSetName = "ColorBackground")]
    public class NewWordCloudCommand : PSCmdlet
    {

        #region Constants

        private const float FOCUS_WORD_SCALE = 1.3f;
        private const float BLEED_AREA_SCALE = 1.15f;

        #endregion Constants

        #region static members

        // TBD

        #endregion static members

        #region Parameters

        [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "ColorBackground")]
        [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "ColorBackground-Mono")]
        [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "FileBackground")]
        [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "FileBackground-Mono")]
        [Alias("InputString", "Text", "String", "Words", "Document", "Page")]
        [AllowEmptyString()]
        public PSObject InputObject { get; set; }

        private string[] _resolvedPaths;
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "ColorBackground")]
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "ColorBackground-Mono")]
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "FileBackground")]
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "FileBackground-Mono")]
        [Alias("OutFile", "ExportPath", "ImagePath")]
        public string[] Path { get; set; }

        [Parameter]
        [ArgumentCompleter(typeof(ImageSizeCompleter))]
        [TransformToSKSizeI]
        public SKSizeI ImageSize { get; set; } = new SKSizeI(4096, 2304);

        [Parameter]
        [Alias("FontFamily", "FontFace")]
        [ArgumentCompleter(typeof(FontFamilyCompleter))]
        [TransformToSKTypeface]
        public SKTypeface Typeface { get; set; } = WCUtils.FontManager.MatchFamily(
            "Consolas", SKFontStyle.Normal);

        [Parameter]
        [Alias("Backdrop", "CanvasColor")]
        [ArgumentCompleter(typeof(SKColorCompleter))]
        [TransformToSKColor]
        public SKColor BackgroundColor { get; set; } = SKColors.Black;

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
        [ValidateRange(0.01, 5)]
        public float WordScale { get; set; } = 1f;

        [Parameter]
        [Alias("Spacing")]
        public float Padding { get; set; } = 5f;

        [Parameter]
        [ValidateRange(1, 500)]
        public float DistanceStep { get; set; } = 5;

        [Parameter]
        [ValidateRange(1, 50)]
        public float RadialStep { get; set; } = 15;

        [Parameter]
        [Alias("MaxWords")]
        [ValidateRange(0, 1000)]
        public ushort MaxRenderedWords { get; set; } = 100;

        [Parameter]
        [Alias("SeedValue")]
        public int RandomSeed { get; set; }

        [Parameter]
        [Alias()]
        public SwitchParameter DisableRotation { get; set; }

        [Parameter]
        public SwitchParameter PassThru { get; set; }

        #endregion Parameters

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
            "yours","yourself","yourselves",
        };

        private static readonly char[] _splitChars = new[] {
            ' ','.',',','"','?','!','{','}','[',']',':','(',')','“','”','*','#','%','^','&','+','=' };

        private List<Task<IEnumerable<string>>> _wordProcessingTasks;
        private float _fontScale;
        private float _wordScaleRange;
        private Random _random;
        private int _progressID;
        private int _colorIndex = 0;
        private SKColor _nextColor
        {
            get
            {
                if (_colorIndex >= ColorSet.Length)
                {
                    _colorIndex = 0;
                }

                var color = ColorSet[_colorIndex];
                _colorIndex++;

                return color;
            }
        }

        protected override void BeginProcessing()
        {
            _random = MyInvocation.BoundParameters.ContainsKey("RandomSeed") ? new Random(RandomSeed) : new Random();
            _progressID = _random.Next();

            var targetPaths = new List<string>();

            foreach (string path in Path)
            {
                var resolvedPaths = SessionState.Path.GetUnresolvedProviderPathFromPSPath(
                    path, out ProviderInfo provider, out PSDriveInfo drive);
                if (resolvedPaths != null)
                {
                    targetPaths.Add(resolvedPaths);
                }
            }

            if (targetPaths.Count == 0)
            {
                ThrowTerminatingError(new ErrorRecord(
                    new ArgumentException("Unable to resolve any recognisable paths", "Path"),
                    "PSWordCloud.BadPath",
                    ErrorCategory.InvalidArgument,
                    Path));
            }

            _resolvedPaths = targetPaths.ToArray();
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
                text = InputObject.BaseObject as string[];
                if (text == null)
                {
                    text = new[] { InputObject.BaseObject as string };
                }
            }

            if (_wordProcessingTasks == null)
            {
                _wordProcessingTasks = new List<Task<IEnumerable<string>>>(text.Length);
            }

            foreach (var line in text)
            {
                _wordProcessingTasks.Add(Task.Run(async () => await ProcessLineAsync(line)));
            }
        }

        protected override void EndProcessing()
        {
            var lineStrings = Task.WhenAll<IEnumerable<string>>(_wordProcessingTasks);
            lineStrings.Wait();

            var wordCount = 0;
            var wordScaleDictionary = new Dictionary<string, float>(StringComparer.OrdinalIgnoreCase);
            SKRect wordBounds = SKRect.Empty;
            SKPath wordPath = null;
            SKRectI drawableBounds = new SKRectI(0, 0, ImageSize.Width, ImageSize.Height);
            SKRegion clipRegion = null;

            float inflationValue = 0;

            foreach (var lineWords in lineStrings.Result)
            {
                CountWords(lineWords, wordScaleDictionary);
            }

            // All words counted and in the dictionary.
            var highestWordFreq = wordScaleDictionary.Values.Max();

            if (FocusWord != null)
            {
                wordScaleDictionary[FocusWord] = highestWordFreq = highestWordFreq * FOCUS_WORD_SCALE;
            }

            _wordScaleRange = highestWordFreq - wordScaleDictionary.Values.Min();
            float averageWordFrequency = wordScaleDictionary.Values.Average();

            List<string> sortedWordList = new List<string>(SortWordList(wordScaleDictionary));

            try
            {
                wordPath = new SKPath();
                clipRegion = new SKRegion();
                clipRegion.SetRect(drawableBounds);

                SetFontScale(clipRegion, averageWordFrequency, wordScaleDictionary.Count);

                var scaledWordSizes = new Dictionary<string, float>(
                    sortedWordList.Count, StringComparer.OrdinalIgnoreCase);

                bool retry;
                using (SKPaint brush = new SKPaint())
                {
                    brush.Typeface = Typeface;
                    do
                    {
                        retry = false;
                        foreach (string word in sortedWordList)
                        {
                            var adjustedWordSize = ScaleWordSize(wordScaleDictionary[word], wordScaleDictionary);

                            // If the final word size is too small, it probably won't be visible in the final image anyway
                            if (adjustedWordSize < 5) continue;

                            brush.TextSize = adjustedWordSize;
                            var adjustedTextWidth = brush.MeasureText(word) + Padding;

                            if ((DisableRotation.IsPresent && adjustedTextWidth > drawableBounds.Width)
                                || adjustedTextWidth > Math.Max(drawableBounds.Width, drawableBounds.Height))
                            {
                                retry = true;
                                _fontScale *= 0.98f;
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

                using (SKFileWStream streamWriter = new SKFileWStream(_resolvedPaths[0]))
                using (SKXmlStreamWriter xmlWriter = new SKXmlStreamWriter(streamWriter))
                using (SKCanvas canvas = SKSvgCanvas.Create(drawableBounds, xmlWriter))
                using (SKPaint brush = new SKPaint())
                using (SKRegion occupiedSpace = new SKRegion())
                {
                    if (BackgroundColor != SKColor.Empty)
                    {
                        canvas.Clear(BackgroundColor);
                    }

                    WordOrientation targetOrientation;

                    SKPoint targetPoint;

                    brush.IsAutohinted = true;
                    brush.IsAntialias = true;
                    brush.Style = SKPaintStyle.StrokeAndFill;
                    brush.Typeface = Typeface;

                    var progress = new ProgressRecord(
                        _progressID,
                        "Starting...",
                        "Finding available space to draw...");

                    foreach (string word in sortedWordList.OrderByDescending(x => scaledWordSizes[x]))
                    {
                        wordCount++;
                        wordPath.Reset();

                        inflationValue = brush.StrokeWidth + Padding * scaledWordSizes[word] / 8;
                        targetOrientation = WordOrientation.Horizontal;
                        targetPoint = SKPoint.Empty;

                        brush.NextWord(scaledWordSizes[word], StrokeWidth, _nextColor);

                        progress.Activity = string.Format("Drawing '{0}' at {1:0} em", word, brush.TextSize);
                        progress.PercentComplete = 100 * wordCount / scaledWordSizes.Count;
                        WriteProgress(progress);

                        for (float radius = 0; radius <= maxRadius; radius += GetRadiusIncrement(scaledWordSizes[word]))
                        {
                            brush.MeasureText(word, ref wordBounds);
                            wordBounds.Inflate(inflationValue, inflationValue);
                            SKSize inflatedWordSize = wordBounds.Size;
                            SKPoint adjustedPoint, pointOffset;

                            foreach (var point in GetRadialPoints(centrePoint, radius, aspectRatio))
                            {
                                if (!canvas.LocalClipBounds.Contains(point))
                                {
                                    continue;
                                }

                                SKMatrix matrix = SKMatrix.MakeIdentity();
                                foreach (var orientation in GetRotationModes())
                                {
                                    switch (orientation)
                                    {
                                        case WordOrientation.Vertical:
                                            pointOffset = new SKPoint(
                                                (inflatedWordSize.Height / 2) + (float)(_random.NextDouble() + 0.25),
                                                (inflatedWordSize.Width / 2) + (float)(_random.NextDouble() + 0.25));
                                            SKMatrix.RotateDegrees(ref matrix, 90, centrePoint.X, centrePoint.Y);
                                            break;

                                        case WordOrientation.VerticalFlipped:
                                            pointOffset = new SKPoint(
                                                (inflatedWordSize.Height / 2) + (float)(_random.NextDouble() + 0.25),
                                                (inflatedWordSize.Width / 2) + (float)(_random.NextDouble() + 0.25));
                                            SKMatrix.RotateDegrees(ref matrix, -90, centrePoint.X, centrePoint.Y);
                                            break;

                                        default:
                                            pointOffset = new SKPoint(
                                                (inflatedWordSize.Width / 2) + (float)(_random.NextDouble() + 0.25),
                                                (inflatedWordSize.Height / 2) + (float)(_random.NextDouble() + 0.25));
                                            break;
                                    }

                                    adjustedPoint = SKPoint.Subtract(point, pointOffset);
                                    wordPath = brush.GetTextPath(word, adjustedPoint.X, adjustedPoint.Y);
                                    wordPath.Transform(matrix);
                                    wordPath.GetBounds(out SKRect bounds);
                                    if (!occupiedSpace.IntersectsRect(bounds) && !bounds.FallsOutside(clipRegion))
                                    {
                                        targetPoint = adjustedPoint;
                                        targetOrientation = orientation;
                                        goto nextWord;
                                    }
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
                            if (targetOrientation == WordOrientation.Vertical)
                            {
                                brush.IsVerticalText = true;
                            }

                            occupiedSpace.Op(wordPath, SKRegionOperation.Union);
                            canvas.DrawPath(wordPath, brush);

                            if (MyInvocation.BoundParameters.ContainsKey("StrokeWidth"))
                            {
                                brush.Color = StrokeColor;
                                brush.IsStroke = true;
                                canvas.DrawPath(wordPath, brush);
                            }
                        }
                    }

                    canvas.Flush();
                    streamWriter.Flush();
                    var file = InvokeProvider.Item.Get(_resolvedPaths[0]);
                    if (PassThru.IsPresent)
                    {
                        WriteObject(file, true);
                    }

                    if (_resolvedPaths.Length > 1)
                    {
                        foreach (string path in _resolvedPaths)
                        {
                            if (path == _resolvedPaths[0]) continue;
                            InvokeProvider.Item.Copy(
                                _resolvedPaths[0], path, false,
                                CopyContainers.CopyTargetContainer);
                            WriteObject(InvokeProvider.Item.Get(path), true);
                        }
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
            }
        }

        private float GetRadiusIncrement(float wordSize)
            => ((float)_random.NextDouble() * wordSize * DistanceStep) / Math.Max(1, 21 - Padding * 2);

        private void CountWords(IEnumerable<string> wordList, IDictionary<string, float> dictionary)
        {
            foreach (string word in wordList)
            {
                var trimmedWord = System.Text.RegularExpressions.Regex.Replace(
                    word, "s$", string.Empty, RegexOptions.IgnoreCase);
                var pluralWord = String.Format("{0}s", word);
                if (dictionary.ContainsKey(trimmedWord))
                {
                    dictionary[trimmedWord]++;
                }
                else if (dictionary.ContainsKey(pluralWord))
                {
                    dictionary[word] = dictionary[pluralWord] + 1;
                    dictionary.Remove(pluralWord);
                }
                else if (dictionary.ContainsKey(word))
                {
                    dictionary[word]++;
                }
                else
                {
                    dictionary.Add(word, 1);
                }
            }
        }

        private IEnumerable<WordOrientation> GetRotationModes()
        {
            if (DisableRotation)
            {
                yield return WordOrientation.Horizontal;
                yield break;
            }

            var set = _random.Next() % 4;

            switch (set)
            {
                case 0:
                    yield return WordOrientation.Horizontal;
                    yield return WordOrientation.Vertical;
                    yield break;

                case 1:
                    yield return WordOrientation.Vertical;
                    yield return WordOrientation.Horizontal;
                    yield break;

                case 2:
                    yield return WordOrientation.Horizontal;
                    yield return WordOrientation.VerticalFlipped;
                    yield break;

                case 3:
                    yield return WordOrientation.VerticalFlipped;
                    yield return WordOrientation.Horizontal;
                    yield break;
            }
        }

        private void SetFontScale(SKRegion space, float averageWordFrequency, int wordCount)
        {
            _fontScale = 2.5f * WordScale * (space.Bounds.Height + space.Bounds.Width)
                / (1.5f * averageWordFrequency * Math.Min(wordCount, MaxRenderedWords));
        }

        private float ScaleWordSize(float baseSize, IDictionary<string, float> scaleDictionary)
        {
            float wordDensity
                = scaleDictionary.Values.Max()
                - scaleDictionary.Values.Min()
                - scaleDictionary.Values.Average();

            float scaledSize = baseSize * _fontScale * (0.9f + 2 * (float)_random.NextDouble() / (1 + wordDensity));
            return scaledSize < 0.5f ? scaledSize + 0.5f : scaledSize;
        }

        private IEnumerable<string> SortWordList(IDictionary<string, float> dictionary)
        {
            return dictionary.Keys.OrderByDescending(size => dictionary[size])
                .Take(MaxRenderedWords == 0 ? ushort.MaxValue : MaxRenderedWords);
        }

        private IEnumerable<SKPoint> GetRadialPoints(SKPoint centre, float radius, float aspectRatio = 1)
        {
            if (radius == 0)
            {
                yield return centre;
                yield break;
            }

            Complex point;
            float angle = 0, maxAngle = 0, angleIncrement = 360f / (radius * RadialStep + 1);
            bool clockwise = _random.NextDouble() > 0.5;

            switch (_random.Next() % 4)
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
                point = Complex.FromPolarCoordinates(radius, angle);
                yield return new SKPoint(centre.X + (float)point.Real * aspectRatio, centre.Y + (float)point.Imaginary);

                angle += angleIncrement;
            } while (clockwise ? angle <= maxAngle : angle >= maxAngle);
        }

        private async Task<IEnumerable<string>> ProcessLineAsync(string line)
        {
            return await Task.Run<IEnumerable<string>>(
                () =>
                {
                    var words = new List<string>(line.Split(_splitChars, StringSplitOptions.RemoveEmptyEntries));
                    words.RemoveAll(x =>
                        _stopWords.Contains(x, StringComparer.OrdinalIgnoreCase)
                        || Regex.IsMatch(x, "^[^a-z]+$", RegexOptions.IgnoreCase)
                        || x.Length < 2);
                    return words;
                });
        }
    }
}
