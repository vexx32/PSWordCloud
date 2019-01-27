using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Numerics;
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
        [Alias("FontFamily", "FontFace", "Typeface")]
        [ArgumentCompleter(typeof(FontFamilyCompleter))]
        [TransformToSKTypeface]
        public SKTypeface Font { get; set; } = WCUtils.FontManager.MatchFamily(
            "Consolas", SKFontStyle.Normal);

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
        public float Padding { get; set; } = 3.5f;

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
        [Alias("AllowOverflow")]
        public SwitchParameter AllowBleed { get; set; }

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
            ' ','.',',','"','?','!','{','}','[',']',':','(',')','“','”','*','#','%','^','&','+','='
        };

        private List<Task<string[]>> _wordProcessingTasks;
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
                _wordProcessingTasks = new List<Task<string[]>>(text.Length);
            }

            foreach (var line in text)
            {
                _wordProcessingTasks.Add(Task.Run(async () => await ProcessLineAsync(line)));
            }
        }

        protected override void EndProcessing()
        {
            var lineStrings = Task.WhenAll<string[]>(_wordProcessingTasks);
            lineStrings.Wait();

            var wordCount = 0;
            float initialAngle = 0, angleIncrement = 0, inflationValue = 0;
            SKPath wordPath = null;
            SKRegion wordRegion = null;
            SKRegion drawableClip = null;
            SKRect wordBounds = SKRect.Empty;
            Dictionary<string, float> wordScaleDictionary = new Dictionary<string, float>(
                StringComparer.OrdinalIgnoreCase);

            foreach (var lineWords in lineStrings.Result)
            {
                foreach (string word in lineWords)
                {
                    var trimmedWord = word.TrimEnd('s');
                    var pluralWord = String.Format("{0}s", word);
                    if (wordScaleDictionary.ContainsKey(trimmedWord))
                    {
                        wordScaleDictionary[trimmedWord]++;
                    }
                    else if (wordScaleDictionary.ContainsKey(pluralWord))
                    {
                        wordScaleDictionary[word] = wordScaleDictionary[pluralWord] + 1;
                        wordScaleDictionary.Remove(pluralWord);
                    }
                    else if (wordScaleDictionary.ContainsKey(word))
                    {
                        wordScaleDictionary[word]++;
                    }
                    else
                    {
                        wordScaleDictionary.Add(word, 1);
                    }
                }
            }

            // All words counted and in the dictionary.
            var wordSizeValues = wordScaleDictionary.Values;
            var highestWordFreq = wordSizeValues.Max();
            if (FocusWord != null)
            {
                wordScaleDictionary[FocusWord] = highestWordFreq = highestWordFreq * FOCUS_WORD_SCALE;
            }

            float averageWordFrequency = wordSizeValues.Average();

            List<string> sortedWordList = new List<string>(wordScaleDictionary.Keys
                .OrderByDescending(size => wordScaleDictionary[size])
                .Take(MaxRenderedWords == 0 ? ushort.MaxValue : MaxRenderedWords));

            try
            {
                SKRectI drawableBounds = new SKRectI(0, 0, ImageSize.Width, ImageSize.Height);
                drawableClip = new SKRegion();
                if (AllowBleed.IsPresent)
                {
                    drawableBounds.Inflate(
                        (int)Math.Round(drawableBounds.Width * BLEED_AREA_SCALE),
                        (int)Math.Round(drawableBounds.Height * BLEED_AREA_SCALE));
                }

                drawableClip.SetRect(drawableBounds);

                float fontScale = WordScale * 1.6f *
                        (drawableBounds.Height + drawableBounds.Width) / (averageWordFrequency * sortedWordList.Count);

                Dictionary<string, float> finalWordEmSizes = new Dictionary<string, float>(
                    sortedWordList.Count, StringComparer.OrdinalIgnoreCase);

                using (SKPaint brush = new SKPaint())
                {
                    brush.Typeface = Font;
                    bool retry = false;
                    do
                    {
                        foreach (string word in sortedWordList)
                        {
                            var adjustedWordSize = (float)Math.Round(
                                2 * wordScaleDictionary[word] * fontScale * _random.NextDouble() /
                                (1f + highestWordFreq - wordSizeValues.Min()) + 0.9);

                            // If the final word size is too small, it probably won't be visible in the final image anyway
                            if (adjustedWordSize < 5) continue;

                            brush.TextSize = adjustedWordSize;
                            var adjustedTextWidth = brush.MeasureText(word) * Padding;

                            if (DisableRotation.IsPresent && adjustedTextWidth > drawableBounds.Width)
                            {
                                retry = true;
                                fontScale *= 0.98f;
                                finalWordEmSizes.Clear();
                                break;
                            }
                            else if (adjustedTextWidth > Math.Max(drawableBounds.Width, drawableBounds.Height))
                            {
                                retry = true;
                                fontScale *= 0.98f;
                                finalWordEmSizes.Clear();
                                break;
                            }

                            finalWordEmSizes[word] = adjustedWordSize;
                        }
                    }
                    while (retry);
                }

                var aspectRatio = drawableBounds.Width / (float)drawableBounds.Height;
                SKPoint centrePoint = new SKPoint(drawableBounds.MidX, drawableBounds.MidY);

                // Remove all words that were cut from the final rendering list
                sortedWordList.RemoveAll(x => !finalWordEmSizes.ContainsKey(x));

                using (SKRegion occupiedRegion = new SKRegion())
                {
                    var maxRadialDistance = Math.Max(drawableBounds.Width, drawableBounds.Height) / 2f;

                    using (SKPaint brush = new SKPaint())
                    using (SKFileWStream streamWriter = new SKFileWStream(_resolvedPaths[0]))
                    using (SKXmlStreamWriter xmlWriter = new SKXmlStreamWriter(streamWriter))
                    using (SKCanvas canvas = SKSvgCanvas.Create(drawableBounds, xmlWriter))
                    {
                        wordPath = new SKPath();
                        brush.IsAutohinted = true;
                        brush.IsAntialias = true;

                        foreach (string word in sortedWordList)
                        {
                            wordCount++;
                            wordPath.Reset();
                            brush.TextSize = finalWordEmSizes[word];
                            brush.StrokeWidth = finalWordEmSizes[word] * StrokeWidth / 100;
                            brush.IsStroke = false;
                            brush.Color = _nextColor;
                            inflationValue = brush.StrokeWidth + Padding * finalWordEmSizes[word] / 10;

                            WriteProgress(
                                new ProgressRecord(
                                    _progressID,
                                    string.Format("Drawing '{0}' at {1} em", word, brush.TextSize),
                                    "Finding available space to draw..."));

                            for (float radialDistance = 0;
                                radialDistance <= maxRadialDistance;
                                radialDistance += (float)_random.NextDouble() * finalWordEmSizes[word] * DistanceStep /
                                Math.Max(1, 21 - Padding * 2))
                            {
                                angleIncrement = 3600f / ((radialDistance + 1) * RadialStep);
                                ScanDirection direction = _random.Next() % 2 == 0 ?
                                    ScanDirection.ClockWise : ScanDirection.CounterClockwise;
                                switch (_random.Next() % 4)
                                {
                                    case 0:
                                        initialAngle = 0;
                                        break;
                                    case 1:
                                        initialAngle = 90;
                                        break;
                                    case 2:
                                        initialAngle = 180;
                                        break;
                                    case 3:
                                        initialAngle = 270;
                                        break;
                                }

                                brush.MeasureText(word, ref wordBounds);
                                SKSize inflatedWordSize = wordBounds.Size + new SKSize(inflationValue, inflationValue);
                                if (TryGetAvailableRadialLocation(
                                    centrePoint, radialDistance, direction, initialAngle,
                                    angleIncrement, aspectRatio, inflatedWordSize,
                                    drawableBounds, occupiedRegion, !DisableRotation.IsPresent,
                                    out SKPoint point, out WordOrientation rotation))
                                {
                                    wordPath = brush.GetTextPath(word, point.X, point.Y);
                                    wordRegion = new SKRegion();
                                    wordRegion.SetPath(wordPath, drawableClip);
                                    occupiedRegion.Op(wordRegion, SKRegionOperation.Union);
                                    canvas.DrawPath(wordPath, brush);

                                    if (MyInvocation.BoundParameters.ContainsKey("StrokeWidth"))
                                    {
                                        brush.IsStroke = true;
                                        brush.Color = StrokeColor;
                                        canvas.DrawPath(wordPath, brush);
                                    }

                                    // Return to word loop for next word or end
                                    break;
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
            }
            catch (Exception e)
            {
                WriteError(new ErrorRecord(e, "PSWordCloudError", ErrorCategory.InvalidResult, null));
            }
            finally
            {
                wordPath?.Dispose();
                wordRegion?.Dispose();
                drawableClip?.Dispose();
            }
        }

        private bool TryGetAvailableRadialLocation(
            SKPoint centre,
            float distance,
            ScanDirection direction,
            float initialAngle,
            float angleIncrement,
            float aspectRatio,
            SKSize wordSize,
            SKRect drawableBounds,
            SKRegion occupiedRegion,
            bool allowRotation,
            out SKPoint location,
            out WordOrientation orientation)
        {
            if (_random == null)
            {
                throw new NullReferenceException("Field _random has not been defined");
            }

            location = SKPoint.Empty;
            orientation = WordOrientation.Horizontal;

            bool clockwise = direction == ScanDirection.ClockWise;
            float angle = initialAngle;
            float maxAngle = clockwise ? initialAngle + 360 : initialAngle - 360;
            SKRectI rect;
            var availableOrientations = allowRotation ?
                new[] { WordOrientation.Horizontal, WordOrientation.Vertical } : new[] { WordOrientation.Horizontal };
            var rotatedWordSize = new SKSize(wordSize.Height, wordSize.Width);
            if (direction == ScanDirection.CounterClockwise)
            {
                angleIncrement *= -1;
            }

            do
            {
                var offsetX = 0f;
                var offsetY = 0f;
                var size = SKSize.Empty;
                Complex complex = Complex.FromPolarCoordinates(distance, angle.ToRadians());
                foreach (WordOrientation currentOrientation in availableOrientations)
                {
                    if (currentOrientation == WordOrientation.Vertical)
                    {
                        offsetX = wordSize.Height * 0.5f * (float)(_random.NextDouble() + 0.25);
                        offsetY = wordSize.Width * 0.5f * (float)(_random.NextDouble() + 0.25);
                        size = rotatedWordSize;
                    }
                    else
                    {
                        offsetX = wordSize.Width * 0.5f * (float)(_random.NextDouble() + 0.25);
                        offsetY = wordSize.Height * 0.5f * (float)(_random.NextDouble() + 0.25);
                        size = wordSize;
                    }

                    SKPoint point = new SKPoint(
                            (float)complex.Real * aspectRatio + centre.X - offsetX,
                            (float)complex.Imaginary + centre.Y - offsetY);

                    if (!drawableBounds.Contains(point))
                    {
                        continue;
                    }

                    rect = SKRect.Create(point, size).ToSKRectI();

                    if (!occupiedRegion.Intersects(rect))
                    {
                        location = point;
                        orientation = currentOrientation;
                        return true;
                    }
                }

                angle += angleIncrement;
            } while (clockwise ? angle <= maxAngle : angle >= maxAngle);

            return false;
        }

        private async Task<string[]> ProcessLineAsync(string line)
        {
            return await Task.Run<string[]>(() =>
            {
                var words = new List<string>(line.Split(_splitChars, StringSplitOptions.RemoveEmptyEntries));
                words.RemoveAll(x => Array.IndexOf(_stopWords, x) == -1);
                return words.ToArray();
            });
        }
    }
}
