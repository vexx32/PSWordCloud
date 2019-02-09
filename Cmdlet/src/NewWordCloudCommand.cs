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

        private const float FOCUS_WORD_SCALE = 1.5f;
        private const float BLEED_AREA_SCALE = 1.15f;
        private const float MIN_SATURATION_VALUE = 5f;
        private const float MIN_BRIGHTNESS_DISTANCE = 25f;

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

        private string _resolvedPath;
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
        [Parameter(Mandatory = true, ParameterSetName = "FileBackground")]
        [Parameter(Mandatory = true, ParameterSetName = "FileBackground-Mono")]
        public string BackgroundImage { get; set; }

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

        private IEnumerable<SKColor> _colors;
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
        [Alias()]
        public SwitchParameter DisableRotation { get; set; }

        [Parameter(Mandatory = true, ParameterSetName = "FileBackground-Mono")]
        [Parameter(Mandatory = true, ParameterSetName = "ColorBackground-Mono")]
        [Alias("BlackAndWhite", "Greyscale")]
        public SwitchParameter Monochrome { get; set; }

        [Parameter]
        [Alias()]
        public SwitchParameter AllowStopWords { get; set; }

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
            ' ','\n','\t','\r','.',',',';','\\','/','|',
            ':','"','?','!','{','}','[',']',':','(',')',
            '<','>','“','”','*','#','%','^','&','+','=' };

        private List<Task<IEnumerable<string>>> _wordProcessingTasks;
        private float _fontScale;
        private static Random _random;
        private static Random Random { get => _random = _random ?? new Random(); }
        private static float RandomFloat { get => (float)Random.NextDouble(); }
        private int _progressID;
        private int _colorIndex = 0;
        private SKColor _nextColor
        {
            get
            {
                if (_colorIndex >= _colors.Count())
                {
                    _colorIndex = 0;
                }

                var color = _colors.ElementAt(_colorIndex);
                _colorIndex++;

                return color;
            }
        }

        protected override void BeginProcessing()
        {
            _random = MyInvocation.BoundParameters.ContainsKey(nameof(RandomSeed))
                ? new Random(RandomSeed)
                : new Random();
            _progressID = Random.Next();

            if (ParameterSetName == "FileBackground" || ParameterSetName == "FileBackground-Mono")
            {
                Environment.CurrentDirectory = SessionState.Path.CurrentFileSystemLocation.Path;
                _backgroundFullPath = System.IO.Path.GetFullPath(BackgroundImage);
            }

            _colors = ProcessColorSet(ColorSet, BackgroundColor, MaxRenderedWords, Monochrome)
                .OrderByDescending(x => x.SortValue(RandomFloat));
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
            SKRectI drawableBounds;
            SKPath wordPath = null;
            SKRegion clipRegion = null;
            SKBitmap backgroundImage = null;

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

            float averageWordFrequency = wordScaleDictionary.Values.Average();

            List<string> sortedWordList = new List<string>(SortWordList(wordScaleDictionary, MaxRenderedWords));

            try
            {
                if (ParameterSetName.StartsWith("FileBackground"))
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
                clipRegion.SetRect(drawableBounds);

                _fontScale = SetFontScale(
                    clipRegion,
                    WordScale, averageWordFrequency,
                    Math.Min(wordScaleDictionary.Count, MaxRenderedWords));

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
                            var adjustedWordSize = ScaleWordSize(wordScaleDictionary[word], _fontScale, wordScaleDictionary);

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
                SKPoint centrePoint = new SKPoint(drawableBounds.Width / 2f, drawableBounds.Height / 2f);

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
                    brush.Style = SKPaintStyle.StrokeAndFill;
                    brush.Typeface = Typeface;

                    var progress = new ProgressRecord(
                        _progressID,
                        "Starting...",
                        "Finding available space to draw...");

                    float radius = 0;
                    foreach (string word in sortedWordList.OrderByDescending(x => scaledWordSizes[x]))
                    {
                        wordCount++;
                        wordPath.Reset();

                        inflationValue = StrokeWidth + Padding * scaledWordSizes[word] * 0.03f;
                        targetOrientation = WordOrientation.Horizontal;
                        targetPoint = SKPoint.Empty;

                        var wordColor = _nextColor;
                        brush.NextWord(scaledWordSizes[word], StrokeWidth, wordColor);

                        progress.Activity = string.Format(
                            "Drawing '{0}' at {1:0} em ({2} of {3})",
                            word, brush.TextSize, wordCount, scaledWordSizes.Count);
                        progress.PercentComplete = 100 * wordCount / scaledWordSizes.Count;
                        WriteProgress(progress);

                        var radialIncrement = GetRadiusIncrement(scaledWordSizes[word], DistanceStep, Padding);
                        for (radius /= 3; radius <= maxRadius; radius += radialIncrement)
                        {
                            brush.MeasureText(word, ref wordBounds);
                            wordBounds.Inflate(inflationValue, inflationValue);
                            SKPoint adjustedPoint, pointOffset;

                            foreach (var point in GetRadialPoints(centrePoint, radius, RadialStep, aspectRatio))
                            {
                                if (!canvas.LocalClipBounds.Contains(point))
                                {
                                    continue;
                                }

                                SKMatrix matrix = SKMatrix.MakeIdentity();
                                foreach (var orientation in GetRotationModes(allowRotation: !DisableRotation))
                                {
                                    switch (orientation)
                                    {
                                        case WordOrientation.Vertical:
                                            pointOffset = new SKPoint(
                                                (wordBounds.Height / 2) + RandomFloat - 0.5f,
                                                (wordBounds.Width / 2) + RandomFloat - 0.5f);
                                            adjustedPoint = point - pointOffset;
                                            SKMatrix.RotateDegrees(ref matrix, 90, adjustedPoint.X, adjustedPoint.Y);
                                            break;

                                        case WordOrientation.VerticalFlipped:
                                            pointOffset = new SKPoint(
                                                (wordBounds.Height / 2) + RandomFloat - 0.5f,
                                                (wordBounds.Width / 2) + RandomFloat - 0.5f);
                                            adjustedPoint = point - pointOffset;
                                            SKMatrix.RotateDegrees(ref matrix, -90, adjustedPoint.X, adjustedPoint.Y);
                                            break;

                                        default:
                                            pointOffset = new SKPoint(
                                                (wordBounds.Width / 2) + RandomFloat - 0.5f,
                                                (wordBounds.Height / 2) + RandomFloat - 0.5f);
                                            adjustedPoint = point - pointOffset;
                                            break;
                                    }

                                    wordPath = brush.GetTextPath(word, adjustedPoint.X, adjustedPoint.Y);
                                    wordPath.Transform(matrix);
                                    wordPath.GetBounds(out wordBounds);
                                    wordPath.FillType = SKPathFillType.Winding;

                                    wordBounds.Inflate(inflationValue, inflationValue);
                                    if (!occupiedSpace.IntersectsRect(wordBounds)
                                        && !wordBounds.FallsOutside(clipRegion))
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

                                radius += radialIncrement;
                            }
                        }

                    nextWord:
                        if (targetPoint != SKPoint.Empty)
                        {
                            if (MyInvocation.BoundParameters.ContainsKey(nameof(StrokeWidth)))
                            {
                                brush.Color = StrokeColor;
                                brush.IsStroke = true;
                                canvas.DrawPath(wordPath, brush);

                                brush.IsStroke = false;
                                brush.Color = wordColor;
                            }

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
            }
        }

        private static IEnumerable<SKColor> ProcessColorSet(
            SKColor[] set, SKColor background, int maxCount, bool monochrome)
        {
            Random.Shuffle(set);
            background.ToHsv(out float bh, out float bs, out float backgroundBrightness);

            foreach (var color in set.Take(maxCount))
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

        private static float GetRadiusIncrement(float wordSize, float distanceStep, float padding)
            => (2 * RandomFloat * wordSize * distanceStep) / Math.Max(1, 21 - padding * 2);

        private static void CountWords(IEnumerable<string> wordList, IDictionary<string, float> dictionary)
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

        private static IEnumerable<WordOrientation> GetRotationModes(bool allowRotation)
        {
            if (allowRotation)
            {
                if (RandomFloat > 0.5)
                {
                    yield return WordOrientation.Horizontal;
                    yield return RandomFloat > 0.5
                        ? WordOrientation.Vertical
                        : WordOrientation.VerticalFlipped;
                }
                else
                {
                    yield return RandomFloat > 0.5
                        ? WordOrientation.Vertical
                        : WordOrientation.VerticalFlipped;
                    yield return WordOrientation.Horizontal;
                }
            }
            else
            {
                yield return WordOrientation.Horizontal;
                yield break;
            }
        }

        private static float SetFontScale(SKRegion space, float baseScale, float averageWordFrequency, int wordCount)
        {
            return baseScale * (space.Bounds.Height + space.Bounds.Width)
                / (1.5f * averageWordFrequency * wordCount);
        }

        private static float ScaleWordSize(
            float baseSize, float globalScale, IDictionary<string, float> scaleDictionary)
        {
            float wordDensity
                = scaleDictionary.Values.Max()
                - scaleDictionary.Values.Min()
                - scaleDictionary.Values.Average();

            float scaledSize = baseSize * globalScale * (0.9f + 3 * RandomFloat / (1 + wordDensity));
            return scaledSize < 0.5f ? scaledSize + 0.5f : scaledSize;
        }

        private static IEnumerable<string> SortWordList(IDictionary<string, float> dictionary, int maxWords)
        {
            return dictionary.Keys.OrderByDescending(word => dictionary[word])
                .Take(maxWords == 0 ? int.MaxValue : maxWords);
        }

        private static IEnumerable<SKPoint> GetRadialPoints(
            SKPoint centre, float radius, float radialStep, float aspectRatio = 1)
        {
            if (radius == 0)
            {
                yield return centre;
                yield break;
            }

            Complex point;
            float angle = 0, maxAngle = 0, angleIncrement = 360f / (radius * radialStep + 1);
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
                point = Complex.FromPolarCoordinates(radius, angle);
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
                    words.RemoveAll(x =>
                        (_stopWords.Contains(x, StringComparer.OrdinalIgnoreCase) && !AllowStopWords)
                        || Regex.IsMatch(x, "^[^a-z]+$", RegexOptions.IgnoreCase)
                        || x.Length < 2);
                    return words;
                });
        }
    }
}
