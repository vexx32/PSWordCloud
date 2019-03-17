using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Numerics;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using SkiaSharp;

[assembly: InternalsVisibleTo("PSWordCloud.Tests")]
namespace PSWordCloud
{
    /// <summary>
    /// Defines the New-WordCloud cmdlet.
    ///
    /// This command can be used to input large amounts of text, and will generate a word cloud based on
    /// the relative frequencies of the words in the input text.
    /// </summary>
    [Cmdlet(VerbsCommon.New, "WordCloud", DefaultParameterSetName = "ColorBackground",
        HelpUri = "https://github.com/vexx32/PSWordCloud/blob/master/docs/New-WordCloud.md")]
    [Alias("wordcloud", "nwc", "wcloud")]
    [OutputType(typeof(System.IO.FileInfo))]
    public class NewWordCloudCommand : PSCmdlet
    {

        #region Constants

        private const float FOCUS_WORD_SCALE = 1.3f;
        private const float BLEED_AREA_SCALE = 1.15f;
        private const float MIN_SATURATION_VALUE = 5f;
        private const float MIN_BRIGHTNESS_DISTANCE = 25f;
        private const float MAX_WORD_WIDTH_PERCENT = 0.75f;
        private const float PADDING_BASE_SCALE = 0.05f;

        internal const float STROKE_BASE_SCALE = 0.02f;

        #endregion Constants

        #region StaticMembers

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

        private static readonly object _randomLock = new object();
        private static Random _random;
        private static Random Random => _random = _random ?? new Random();

        #endregion StaticMembers

        #region Parameters

        /// <summary>
        /// Gets or sets the input text to supply to the word cloud. All input is accepted, but will be treated
        /// as string data regardless of the input type. If you are entering complex object input, ensure they
        /// have a meaningful ToString() method override defined.
        /// </summary>
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
        [Parameter(Mandatory = true, ParameterSetName = "FileBackground")]
        [Parameter(Mandatory = true, ParameterSetName = "FileBackground-Mono")]
        public string BackgroundImage
        {
            get => _backgroundFullPath;
            set
            {
                var resolvedPaths = SessionState.Path.GetResolvedPSPathFromPSPath(value);
                if (resolvedPaths.Count > 1)
                {
                    throw new ArgumentException(string.Format(
                        "Unable to resolve argument for parameter {0} to a single file path.", nameof(BackgroundImage)));
                }

                _backgroundFullPath = resolvedPaths[0].Path;
            }
        }

        /// <summary>
        /// <para>Gets or sets the image size value for the word cloud image.</para>
        /// <para>Input can be passed directly as a SkiaSharp.SKSizeI object, or in one of the following formats:</para>
        /// <para>1. A predefined size string. One of:
        ///      - 720p           (canvas size: 1280x720)
        ///      - 1080p          (canvas size: 1920x1080)
        ///      - 4K             (canvas size: 3840x2160)
        ///      - A4             (canvas size: 816x1056)
        ///      - Poster11x17    (canvas size: 1056x1632)
        ///      - Poster18x24    (canvas size: 1728x2304)
        ///      - Poster24x36    (canvas size: 2304x3456)</para>
        ///
        /// <para>2. Single integer (e.g., -ImageSize 1024). This will be used as both the width and height of the
        /// image, creating a square canvas.</para>
        /// <para>3. Any image size string (e.g., 1024x768). The first number will be used as the width, and the
        /// second number used as the height of the canvas.</para>
        /// <para>4. A hashtable or custom object with keys or properties named "Width" and "Height" that contain
        /// integer values</para>
        /// </summary>
        /// <value>The default value is a size of 3840x2160.</value>
        [Parameter(ParameterSetName = "ColorBackground")]
        [Parameter(ParameterSetName = "ColorBackground-Mono")]
        [ArgumentCompleter(typeof(ImageSizeCompleter))]
        [TransformToSKSizeI()]
        public SKSizeI ImageSize { get; set; } = new SKSizeI(3840, 2160);

        /// <summary>
        /// <para>Gets or sets the typeface to be used in the word cloud.</para>
        /// <para>Input can be processed as a SkiaSharp.SKTypeface object, or one of the following formats:</para>
        /// <para>1. String value matching a valid font name. These can be autocompleted by pressing [Tab].
        /// An invalid value will cause the system default to be used.</para>
        /// <para>2. A custom object or hashtable object containing the following keys or properties:
        ///     - FamilyName: string value. If no font by this name is available, the system default will be used.
        ///     - FontWeight: "Invisible", "ExtraLight", Light", "Thin", "Normal", "Medium", "SemiBold", "Bold",
        ///       "ExtraBold", "Black", "ExtraBlack" (Default: "Normal")
        ///     - FontSlant: "Upright", "Italic", "Oblique" (Default: "Upright")
        ///     - FontWidth: "UltraCondensed", "ExtraCondensed", "Condensed", "SemiCondensed", "Normal", "SemiExpanded",
        ///       "Expanded", "ExtraExpanded", "UltraExpanded" (Default: "Normal")</para>
        /// </summary>
        /// <value>The default value is the font "Consolas" with Normal styles.</value>
        [Parameter()]
        [Alias("FontFamily", "FontFace")]
        [ArgumentCompleter(typeof(FontFamilyCompleter))]
        [TransformToSKTypeface()]
        public SKTypeface Typeface { get; set; } = WCUtils.FontManager.MatchFamily("Consolas", SKFontStyle.Normal);

        /// <summary>
        /// <para>Gets or sets the SKColor value used as the background for the word cloud image.</para>
        /// <para>Accepts input as a complete SKColor object, or one of the following formats:</para>
        /// <para>1. A string color name matching one of the fields in SkiaSharp.SKColors. These values will be pulled
        /// for tab-completion automatically.</para>
        /// <para>2. A hexadecimal number string with or without the preceding #, in the form:
        /// AARRGGBB, RRGGBB, ARGB, or RGB.</para>
        /// <para>3. A hashtable or custom object with keys or properties named: "Red","Green","Blue", and/or "Alpha".
        /// Values may range from 0-255. Omitted color values are assumed to be 0, but omitting alpha defaults it to
        /// 255 (fully opaque).</para>
        /// </summary>
        /// <value>The default value is SKColors.Black.</value>
        [Parameter(ParameterSetName = "ColorBackground")]
        [Parameter(ParameterSetName = "ColorBackground-Mono")]
        [Alias("Backdrop", "CanvasColor")]
        [ArgumentCompleter(typeof(SKColorCompleter))]
        [TransformToSKColor()]
        public SKColor BackgroundColor { get; set; } = SKColors.Black;

        private List<SKColor> _colors;
        /// <summary>
        /// <para>Gets or sets the SKColor values used for the words in the cloud. Multiple values are accepted, and
        /// each new word pulls the next color from the set. If the end of the set is reached, the next word will
        /// reset the index to the start and retrieve the first color again.</para>
        /// <para>Accepts input as a complete SKColor object, or one of the following formats:</para>
        /// <para>1. A string color name matching one of the fields in SkiaSharp.SKColors. These values will be pulled
        /// for tab-completion automatically. Names containing wildcards may be used, and all matching colors will be
        /// included in the set.</para>
        /// <para>2. A hexadecimal number string with or without the preceding #, in the form:
        /// AARRGGBB, RRGGBB, ARGB, or RGB.</para>
        /// <para>3. A hashtable or custom object with keys or properties named: "Red","Green","Blue", and/or "Alpha".
        /// Values may range from 0-255. Omitted color values are assumed to be 0, but omitting alpha defaults it to
        /// 255 (fully opaque).</para>
        /// </summary>
        /// <value>The default value is all available named colors in SkiaSharp.SKColors.</value>
        [Parameter()]
        [SupportsWildcards()]
        [TransformToSKColor()]
        [ArgumentCompleter(typeof(SKColorCompleter))]
        public SKColor[] ColorSet { get; set; } = WCUtils.StandardColors.ToArray();

        /// <summary>
        /// Gets or sets the width of the word outline. Values from 0-10 are permitted. A zero value indicates the
        /// special "Hairline" width, where the width of the stroke depends on the SVG viewing scale.
        /// </summary>
        [Parameter()]
        [Alias("OutlineWidth")]
        [ValidateRange(0, 10)]
        public float StrokeWidth { get; set; }

        /// <summary>
        /// <para>Gets or sets the SKColor value used as the stroke color for the words in the image.</para>
        /// <para>Accepts input as a complete SKColor object, or one of the following formats:</para>
        /// <para>1. A string color name matching one of the fields in SkiaSharp.SKColors. These values will be pulled
        /// for tab-completion automatically.</para>
        /// <para>2. A hexadecimal number string with or without the preceding #, in the form:
        /// AARRGGBB, RRGGBB, ARGB, or RGB.</para>
        /// <para>3. A hashtable or custom object with keys or properties named: "Red","Green","Blue", and/or "Alpha".
        /// Values may range from 0-255. Omitted color values are assumed to be 0, but omitting alpha defaults it to
        /// 255 (fully opaque).</para>
        /// </summary>
        /// <value>The default value is SKColors.Black.</value>
        [Parameter()]
        [Alias("OutlineColor")]
        [TransformToSKColor()]
        [ArgumentCompleter(typeof(SKColorCompleter))]
        public SKColor StrokeColor { get; set; } = SKColors.Black;

        /// <summary>
        /// Gets or sets the focus word string to be used in the word cloud. This string will typically appear in the
        /// centre of the cloud, larger than all the other words.
        /// </summary>
        [Parameter()]
        [Alias("Title")]
        public string FocusWord { get; set; }

        /// <summary>
        /// <para>Gets or sets the words to be explicitly ignored when rendering the word cloud.</para>
        /// <para>This is usually used to exclude irrelevant words, link segments, etc.</para>
        /// </summary>
        [Parameter()]
        [Alias("ForbidWord", "IgnoreWord")]
        public string[] ExcludeWord { get; set; }

        /// <summary>
        /// <para>Gets or sets the words to be explicitly included in rendering of the cloud.</para>
        /// <para>This can be used to override specific words normally excluded by the StopWords list.</para>
        /// </summary>
        /// <value></value>
        [Parameter()]
        [Alias()]
        public string[] IncludeWord { get; set; }

        /// <summary>
        /// Gets or sets the float value to scale the base word size by. By default, the word cloud is scaled to fill
        /// most of the canvas. A value of 0.5 should result in the cloud covering approximately half of the canvas,
        /// clustered around the center.
        /// </summary>
        /// <value>The default value is 1.</value>
        [Parameter()]
        [Alias("ScaleFactor")]
        [ValidateRange(0.01, 20)]
        public float WordScale { get; set; } = 1;

        /// <summary>
        /// Gets or sets the float value to scale the padding space around the words by.
        /// </summary>
        /// <value>The default value is 5.</value>
        [Parameter()]
        [Alias("Spacing")]
        public float Padding { get; set; } = 5;

        /// <summary>
        /// Gets or sets the value to scale the distance step by. Larger numbers will result in more radially spaced
        /// out clouds.
        /// </summary>
        /// <value>The default value is 5.</value>
        [Parameter()]
        [ValidateRange(1, 500)]
        public float DistanceStep { get; set; } = 5;

        /// <summary>
        /// Gets or sets the value to scale the radial arc step by.
        /// </summary>
        /// <value>The default value is 15.</value>
        [Parameter()]
        [ValidateRange(1, 50)]
        public float RadialStep { get; set; } = 15;

        /// <summary>
        /// Gets or sets the maximum number of words to render as part of the cloud.
        /// </summary>
        /// <value>The default value is 100.</value>
        [Parameter()]
        [Alias("MaxWords")]
        [ValidateRange(0, int.MaxValue)]
        public int MaxRenderedWords { get; set; } = 100;

        /// <summary>
        /// Gets or sets the maximum number of colors to use from the values contained in the ColorSet parameter.
        /// The values in the ColorSet parameter are shuffled before being trimmed down here.
        /// </summary>
        /// <value>The default value is int.MaxValue.</value>
        [Parameter()]
        [Alias("MaxColours")]
        [ValidateRange(1, int.MaxValue)]
        public int MaxColors { get; set; } = int.MaxValue;

        /// <summary>
        /// Gets or sets the seed value for the random numbers used to vary the position and placement patterns.
        /// </summary>
        [Parameter()]
        [Alias("SeedValue")]
        public int RandomSeed { get; set; }

        /// <summary>
        /// Gets or sets which types of word rotations are used when drawing the word cloud.
        /// </summary>
        [Parameter()]
        [Alias()]
        public WordOrientations AllowRotation { get; set; } = WordOrientations.EitherVertical;

        /// <summary>
        /// Gets or sets whether to draw the cloud in monochrome (greyscale).
        /// </summary>
        [Parameter(Mandatory = true, ParameterSetName = "FileBackground-Mono")]
        [Parameter(Mandatory = true, ParameterSetName = "ColorBackground-Mono")]
        [Alias("BlackAndWhite", "Greyscale")]
        public SwitchParameter Monochrome { get; set; }

        /// <summary>
        /// Gets or sets whether to allow the "banned" words. The StopWords list is comprised of very commonly-used
        /// words, conjunctions, articles, etc., that would otherwise dominate the cloud without adding any real value.
        /// </summary>
        [Parameter()]
        [Alias("IgnoreStopWords")]
        public SwitchParameter AllowStopWords { get; set; }

        /// <summary>
        /// Gets or sets whether or not to allow words to overflow the base canvas.
        /// </summary>
        /// <value></value>
        [Parameter()]
        [Alias("AllowBleed")]
        public SwitchParameter AllowOverflow { get; set; }

        /// <summary>
        /// Gets or sets the value that determines whether or not to retrieve and output the FileInfo object that
        /// represents the completed word cloud when processing is completed.
        /// </summary>
        [Parameter()]
        public SwitchParameter PassThru { get; set; }

        #endregion Parameters

        #region privateVariables

        private List<Task<IEnumerable<string>>> _wordProcessingTasks;

        private float _fontScale;

        private int _progressID;

        private int _colorIndex = 0;

        private SKColor GetNextColor()
        {
            if (_colorIndex == _colors.Count)
            {
                _colorIndex = 0;
            }

            var color = _colors[_colorIndex];
            _colorIndex++;

            return color;
        }

        private float NextDrawAngle()
        {
            switch (AllowRotation)
            {
                case WordOrientations.Vertical:
                    return RandomFloat() > 0.5 ? 0 : 90;
                case WordOrientations.FlippedVertical:
                    return RandomFloat() < 0.5 ? 0 : -90;
                case WordOrientations.EitherVertical:
                    return RandomFloat() < 0.5
                        ? 0
                        : RandomFloat() > 0.5
                            ? 90
                            : -90;
                case WordOrientations.UprightDiagonals:
                    switch (RandomInt(0, 5))
                    {
                        case 0: return -90;
                        case 1: return -45;
                        case 2: return 45;
                        case 3: return 90;
                        case 4: default: return 0;
                    }
                case WordOrientations.InvertedDiagonals:
                    switch (RandomInt(0, 5))
                    {
                        case 0: return 90;
                        case 1: return 135;
                        case 2: return -135;
                        case 3: return -90;
                        case 4: default: return 180;
                    }
                case WordOrientations.AllDiagonals:
                    switch (RandomInt(0, 8))
                    {
                        case 0: return 45;
                        case 1: return 90;
                        case 2: return 135;
                        case 3: return 180;
                        case 4: return -135;
                        case 5: return -90;
                        case 6: return -45;
                        case 7: default: return 0;
                    }
                case WordOrientations.AllUpright:
                    return RandomInt(-90, 91);
                case WordOrientations.AllInverted:
                    return RandomInt(90, 271);
                case WordOrientations.All:
                    return RandomInt(0, 361);
                default:
                    return 0;
            }
        }

        private float _paddingMultiplier => Padding * PADDING_BASE_SCALE;

        #endregion privateVariables

        /// <summary>
        /// Implements the BeginProcessing method for New-WordCloud.
        /// Instantiates the random number generator, and organises the base color set for the cloud.
        /// </summary>
        protected override void BeginProcessing()
        {
            lock (_randomLock)
            {
                _random = MyInvocation.BoundParameters.ContainsKey(nameof(RandomSeed))
                    ? new Random(RandomSeed)
                    : new Random();
            }

            _progressID = RandomInt();

            _colors = ProcessColorSet(ColorSet, BackgroundColor, StrokeColor, MaxRenderedWords, Monochrome)
                .OrderByDescending(x => x.SortValue(RandomFloat()))
                .ToList();
        }

        /// <summary>
        /// Implements the ProcessRecord method for PSWordCloud.
        /// Spins up a Task&lt;IEnumerable&lt;string&gt;&gt; for each input text string to split them all
        /// asynchronously.
        /// </summary>
        protected override void ProcessRecord()
        {
            var text = MyInvocation.ExpectingInput
                    ? new[] { InputObject.BaseObject as string }
                    : InputObject.BaseObject as string[] ?? new[] { InputObject.BaseObject as string };

            _wordProcessingTasks = _wordProcessingTasks ?? new List<Task<IEnumerable<string>>>(text.Length);

            foreach (var line in text)
            {
                _wordProcessingTasks.Add(ProcessInputAsync(line, IncludeWord, ExcludeWord));
            }
        }

        /// <summary>
        /// Implements the EndProcessing method for New-WordCloud.
        /// The majority of the word cloud drawing occurs here.
        /// </summary>
        protected override void EndProcessing()
        {
            var lineStrings = Task.WhenAll<IEnumerable<string>>(_wordProcessingTasks);
            lineStrings.Wait();

            int wordCount = 0;
            float inflationValue, maxWordWidth, highestWordFreq, aspectRatio, maxRadius;

            var wordScaleDictionary = new Dictionary<string, float>(StringComparer.OrdinalIgnoreCase);
            Dictionary<string, float> scaledWordSizes;
            List<string> sortedWordList;

            SKPath wordPath = null;
            SKRegion clipRegion = null;
            SKBitmap backgroundImage = null;
            SKPoint centrePoint;

            SKRect wordBounds = SKRect.Empty, drawableBounds = SKRect.Empty;
            ProgressRecord wordProgress = null, pointProgress = null;

            foreach (var lineWords in lineStrings.Result)
            {
                CountWords(lineWords, wordScaleDictionary);
            }

            // All words counted and in the dictionary.
            highestWordFreq = wordScaleDictionary.Values.Max();

            if (MyInvocation.BoundParameters.ContainsKey(nameof(FocusWord)))
            {
                wordScaleDictionary[FocusWord] = highestWordFreq = highestWordFreq * FOCUS_WORD_SCALE;
            }

            sortedWordList = new List<string>(SortWordList(wordScaleDictionary, MaxRenderedWords));

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

                scaledWordSizes = new Dictionary<string, float>(
                    sortedWordList.Count, StringComparer.OrdinalIgnoreCase);

                maxWordWidth = AllowRotation == WordOrientations.None
                    ? drawableBounds.Width * MAX_WORD_WIDTH_PERCENT
                    : Math.Max(drawableBounds.Width, drawableBounds.Height) * MAX_WORD_WIDTH_PERCENT;

                using (SKPaint brush = new SKPaint())
                {
                    brush.Typeface = Typeface;
                    SKRect rect = SKRect.Empty;
                    float adjustedWordSize;
                    bool retry;

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

                aspectRatio = drawableBounds.Width / (float)drawableBounds.Height;
                centrePoint = new SKPoint(drawableBounds.MidX, drawableBounds.MidY);

                // Remove all words that were cut from the final rendering list
                sortedWordList.RemoveAll(x => !scaledWordSizes.ContainsKey(x));

                maxRadius = Math.Max(drawableBounds.Width, drawableBounds.Height) / 2f;

                using (SKFileWStream outputStream = new SKFileWStream(_resolvedPath))
                using (SKXmlStreamWriter xmlWriter = new SKXmlStreamWriter(outputStream))
                using (SKCanvas canvas = SKSvgCanvas.Create(drawableBounds, xmlWriter))
                using (SKPaint brush = new SKPaint())
                using (SKRegion occupiedSpace = new SKRegion())
                {
                    if (MyInvocation.BoundParameters.ContainsKey(nameof(AllowOverflow)))
                    {
                        drawableBounds.Inflate(
                            drawableBounds.Width * BLEED_AREA_SCALE,
                            drawableBounds.Height * BLEED_AREA_SCALE);
                    }

                    if (ParameterSetName.StartsWith("FileBackground"))
                    {
                        canvas.DrawBitmap(backgroundImage, 0, 0);
                    }
                    else if (BackgroundColor != SKColor.Empty)
                    {
                        canvas.Clear(BackgroundColor);
                    }

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
                        targetPoint = SKPoint.Empty;

                        var wordColor = GetNextColor();
                        brush.NextWord(scaledWordSizes[word], StrokeWidth, wordColor);

                        wordPath.Dispose();
                        wordPath = brush.GetTextPath(word, 0, 0);
                        wordBounds = wordPath.ComputeTightBounds();

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

                                var drawAngle = NextDrawAngle();
                                pointProgress.Activity = string.Format(
                                    "Finding available space to draw at angle: {0}",
                                    drawAngle);
                                pointProgress.StatusDescription = string.Format(
                                    "Checking [Point:{0,8:N2}, {1,8:N2}] ({2,4} / {3,4}) at [Radius: {4,8:N2}]",
                                    point.X, point.Y, pointsChecked, totalPoints, radius);
                                //pointProgress.PercentComplete = 100 * pointsChecked / totalPoints;
                                WriteProgress(pointProgress);

                                baseOffset = new SKPoint(
                                    -(wordWidth / 2),
                                    (wordHeight / 2));
                                adjustedPoint = point + baseOffset;

                                SKMatrix rotation = SKMatrix.MakeRotationDegrees(drawAngle, point.X, point.Y);

                                SKPath alteredPath = brush.GetTextPath(word, adjustedPoint.X, adjustedPoint.Y);
                                alteredPath.Transform(rotation);
                                alteredPath.GetTightBounds(out wordBounds);

                                wordBounds.Inflate(inflationValue * 2, inflationValue * 2);

                                if (wordCount == 1)
                                {
                                    // First word will always be drawn in the centre.
                                    wordPath = alteredPath;
                                    targetPoint = adjustedPoint;
                                    goto nextWord;
                                }
                                else
                                {
                                    if (wordBounds.FallsOutside(clipRegion))
                                    {
                                        continue;
                                    }

                                    if (!occupiedSpace.IntersectsRect(wordBounds))
                                    {
                                        wordPath = alteredPath;
                                        targetPoint = adjustedPoint;
                                        goto nextWord;
                                    }
                                }

                                if (point == centrePoint)
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

        #region HelperMethods

        /// <summary>
        /// Processes the color set to ensure no colors identical to the background or stroke colors are re-used,
        /// and picks out a random selection of colors to use up to the maximum count.
        /// </summary>
        /// <param name="set">The base set of colors to operate on.</param>
        /// <param name="background">The background color, which will be excluded from the final set.</param>
        /// <param name="stroke">The stroke color, which will be excluded from the final set.</param>
        /// <param name="maxCount">The maximum number of colors to return.</param>
        /// <param name="monochrome">If true, indicates to translate all colors to greyscale according to their overall
        /// brightness values.</param>
        /// <returns></returns>
        private static IEnumerable<SKColor> ProcessColorSet(
            SKColor[] set, SKColor background, SKColor stroke, int maxCount, bool monochrome)
        {
            Shuffle(set);
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

        /// <summary>
        /// Counts all given words in the list, and tallies the counts in the given dictionary.
        /// </summary>
        /// <param name="wordList">The input list of words.</param>
        /// <param name="dictionary">The dictionary to tally the counts in.</param>
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

        /// <summary>
        /// Determines the base font scale for the word cloud.
        /// </summary>
        /// <param name="space">The total available drawing space.</param>
        /// <param name="baseScale">The base scale value.</param>
        /// <param name="averageWordFrequency">The average frequency of words.</param>
        /// <param name="wordCount">The total number of words to account for.</param>
        /// <returns>Returns a float value representing a conservative scaling value to apply to each word.</returns>
        private static float FontScale(SKRect space, float baseScale, float averageWordFrequency, int wordCount)
        {
            return baseScale * (space.Height + space.Width)
                / (8 * averageWordFrequency * wordCount);
        }

        /// <summary>
        /// Scale each word by the base word scale value and determine its final font size.
        /// </summary>
        /// <param name="baseSize">The base size for the font.</param>
        /// <param name="globalScale">The global scaling factor.</param>
        /// <param name="scaleDictionary">The dictionary of word scales containing their base sizes.</param>
        /// <returns></returns>
        private static float ScaleWordSize(
            float baseSize, float globalScale, IDictionary<string, float> scaleDictionary)
        {
            return baseSize * globalScale * (2 * RandomFloat()
                / (1 + scaleDictionary.Values.Max() - scaleDictionary.Values.Min()) + 0.9f);
        }

        /// <summary>
        /// Sorts the word list by the frequency of words, in descending order.
        /// </summary>
        /// <param name="dictionary">The dictionary containing words and their relative frequencies.</param>
        /// <param name="maxWords">The total number of words to consider.</param>
        /// <returns>An enumerable string list of words in order from most used to least.</returns>
        private static IEnumerable<string> SortWordList(IDictionary<string, float> dictionary, int maxWords)
        {
            return dictionary.Keys.OrderByDescending(word => dictionary[word])
                .Take(maxWords == 0 ? int.MaxValue : maxWords);
        }

        /// <summary>
        /// Calculates the radius increment to use when scanning for available space to draw.
        /// </summary>
        /// <param name="wordSize">The size of the word currently being drawn.</param>
        /// <param name="distanceStep">The base distance step value.</param>
        /// <param name="maxRadius">The maximum radial distance to scan from the center.</param>
        /// <param name="padding">The padding amount to take into account.</param>
        /// <param name="percentComplete">How close to completion of the cloud we are, and thus how likely it is that
        /// spaces close to the center of the cloud are already filled.</param>
        /// <returns>Returns a float value indicating how far to step along the radius before scanning in a circle
        /// at that radius once again for available space.</returns>
        private static float GetRadiusIncrement(
            float wordSize, float distanceStep, float maxRadius, float padding, float percentComplete)
            => (5 + RandomFloat() * (2.5f + percentComplete / 10)) * distanceStep * wordSize * (1 + padding) / maxRadius;

        /// <summary>
        /// Scans in an ovoid pattern at a given radius to get a set of points to check for sufficient drawing space.
        /// </summary>
        /// <param name="centre">The centre point of the image.</param>
        /// <param name="radius">The current radius we're scanning at.</param>
        /// <param name="radialStep">The current radial stepping value.</param>
        /// <param name="aspectRatio">The aspect ratio of the canvas.</param>
        /// <returns></returns>
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
            bool clockwise = RandomFloat() > 0.5;

            switch (RandomInt() % 4)
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

                angle += angleIncrement * (radialStep / 15);
            } while (clockwise ? angle <= maxAngle : angle >= maxAngle);
        }

        private static float RandomFloat()
        {
            lock (_randomLock)
            {
                return (float)Random.NextDouble();
            }
        }

        private static int RandomInt()
        {
            lock (_randomLock)
            {
                return Random.Next();
            }
        }

        private static int RandomInt(int min, int max)
        {
            lock (_randomLock)
            {
                return Random.Next(min, max);
            }
        }

        private static void Shuffle<T>(T[] items)
        {
            lock (_randomLock)
            {
                Random.Shuffle(items);
            }
        }

        /// <summary>
        /// Asynchronous method used to quickly process large amounts of text input into words.
        /// </summary>
        /// <param name="line">The text to split and process.</param>
        /// <returns>An enumerable string collection of all words in the input, with stopwords stripped out.</returns>
        private Task<IEnumerable<string>> ProcessInputAsync(
            string line, string[] includeWords = null, string[] excludeWords = null)
        {
            return Task.Run<IEnumerable<string>>(
                () =>
                {
                    var words = new List<string>(line.Split(_splitChars, StringSplitOptions.RemoveEmptyEntries));
                    words.RemoveAll(
                        x => includeWords?.Contains(x, StringComparer.OrdinalIgnoreCase) != true
                            && (excludeWords?.Contains(x, StringComparer.OrdinalIgnoreCase) == true
                                || (!AllowStopWords && _stopWords.Contains(x, StringComparer.OrdinalIgnoreCase))
                                || Regex.Replace(x, "[^a-z-]", string.Empty, RegexOptions.IgnoreCase).Length < 2));
                    return words;
                });
        }

        #endregion HelperMethods
    }
}
