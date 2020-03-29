using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Numerics;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
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
    [Cmdlet(VerbsCommon.New, "WordCloud", DefaultParameterSetName = COLOR_BG_SET,
        HelpUri = "https://github.com/vexx32/PSWordCloud/blob/master/docs/New-WordCloud.md")]
    [Alias("wordcloud", "nwc", "wcloud")]
    [OutputType(typeof(System.IO.FileInfo))]
    public class NewWordCloudCommand : PSCmdlet
    {

        #region Constants

        private const float FOCUS_WORD_SCALE = 1.3f;
        private const float BLEED_AREA_SCALE = 1.5f;
        private const float MAX_WORD_WIDTH_PERCENT = 1.0f;
        private const float PADDING_BASE_SCALE = 0.06f;
        private const float MAX_WORD_AREA_PERCENT = 0.0575f;

        private const char ELLIPSIS = '…';

        internal const string COLOR_BG_SET = "ColorBackground";
        internal const string COLOR_BG_FOCUS_SET = "ColorBackground-FocusWord";
        internal const string COLOR_BG_FOCUS_TABLE_SET = "ColorBackground-FocusWord-WordTable";
        internal const string COLOR_BG_TABLE_SET = "ColorBackground-WordTable";
        internal const string FILE_SET = "FileBackground";
        internal const string FILE_FOCUS_SET = "FileBackground-FocusWord";
        internal const string FILE_FOCUS_TABLE_SET = "FileBackground-FocusWord-WordTable";
        internal const string FILE_TABLE_SET = "FileBackground-WordTable";

        internal const float STROKE_BASE_SCALE = 0.01f;

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
        private static Random Random => _random ??= new Random();

        private CancellationTokenSource _cancel = new CancellationTokenSource();

        #endregion StaticMembers

        #region Parameters

        /// <summary>
        /// Gets or sets the input text to supply to the word cloud. All input is accepted, but will be treated
        /// as string data regardless of the input type. If you are entering complex object input, ensure they
        /// have a meaningful ToString() method override defined.
        /// </summary>
        [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = COLOR_BG_SET)]
        [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = COLOR_BG_FOCUS_SET)]
        [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = FILE_SET)]
        [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = FILE_FOCUS_SET)]
        [Alias("InputString", "Text", "String", "Words", "Document", "Page")]
        [AllowEmptyString()]
        public PSObject InputObject { get; set; }

        /// <summary>
        /// Gets or sets the input word dictionary.
        /// Instead of supplying a chunk of text as the input, this parameter allows you to define your own relative
        /// word sizes.
        /// Supply a dictionary or hashtable object where the keys are the words you want to draw in the cloud, and the
        /// values are their relative sizes.
        /// Words will be scaled as a percentage of the largest sized word in the table.
        /// In other words, if you have @{ text = 10; image = 100 }, then "text" will appear 10 times smaller than
        /// "image".
        /// </summary>
        /// <value></value>
        [Parameter(Mandatory = true, ParameterSetName = COLOR_BG_TABLE_SET)]
        [Parameter(Mandatory = true, ParameterSetName = COLOR_BG_FOCUS_TABLE_SET)]
        [Parameter(Mandatory = true, ParameterSetName = FILE_TABLE_SET)]
        [Parameter(Mandatory = true, ParameterSetName = FILE_FOCUS_TABLE_SET)]
        [Alias("WordSizeTable", "CustomWordSizes")]
        public IDictionary WordSizes { get; set; }

        /// <summary>
        /// Gets or sets the output path to save the final SVG vector file to.
        /// </summary>
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = COLOR_BG_SET)]
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = COLOR_BG_FOCUS_SET)]
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = COLOR_BG_FOCUS_TABLE_SET)]
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = COLOR_BG_TABLE_SET)]
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = FILE_SET)]
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = FILE_FOCUS_SET)]
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = FILE_FOCUS_TABLE_SET)]
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = FILE_TABLE_SET)]
        [Alias("OutFile", "ExportPath", "ImagePath")]
        public string Path { get; set; }

        private string _backgroundFullPath;
        /// <summary>
        /// Gets or sets the path to the background image to be used as a base for the final word cloud image.
        /// </summary>
        [Parameter(Mandatory = true, ParameterSetName = FILE_SET)]
        [Parameter(Mandatory = true, ParameterSetName = FILE_FOCUS_SET)]
        [Parameter(Mandatory = true, ParameterSetName = FILE_FOCUS_TABLE_SET)]
        [Parameter(Mandatory = true, ParameterSetName = FILE_TABLE_SET)]
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
        [Parameter(ParameterSetName = COLOR_BG_SET)]
        [Parameter(ParameterSetName = COLOR_BG_FOCUS_SET)]
        [Parameter(ParameterSetName = COLOR_BG_FOCUS_TABLE_SET)]
        [Parameter(ParameterSetName = COLOR_BG_TABLE_SET)]
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
        [Parameter(ParameterSetName = COLOR_BG_SET)]
        [Parameter(ParameterSetName = COLOR_BG_FOCUS_SET)]
        [Parameter(ParameterSetName = COLOR_BG_FOCUS_TABLE_SET)]
        [Parameter(ParameterSetName = COLOR_BG_TABLE_SET)]
        [Alias("Backdrop", "CanvasColor")]
        [ArgumentCompleter(typeof(SKColorCompleter))]
        [TransformToSKColor()]
        public SKColor BackgroundColor { get; set; } = SKColors.Black;

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
        public SKColor[] ColorSet { get; set; } = WCUtils.StandardColors
            .Where(c => c != SKColor.Empty && c.Alpha != 0).ToArray();

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
        [Parameter(Mandatory = true, ParameterSetName = COLOR_BG_FOCUS_SET)]
        [Parameter(Mandatory = true, ParameterSetName = COLOR_BG_FOCUS_TABLE_SET)]
        [Parameter(Mandatory = true, ParameterSetName = FILE_FOCUS_SET)]
        [Parameter(Mandatory = true, ParameterSetName = FILE_FOCUS_TABLE_SET)]
        [Alias("Title")]
        public string FocusWord { get; set; }

        [Parameter(ParameterSetName = COLOR_BG_FOCUS_SET)]
        [Parameter(ParameterSetName = COLOR_BG_FOCUS_TABLE_SET)]
        [Parameter(ParameterSetName = FILE_FOCUS_SET)]
        [Parameter(ParameterSetName = FILE_FOCUS_TABLE_SET)]
        [Alias("RotateTitle", "RotateFocusWord")]
        [ArgumentCompleter(typeof(AngleCompleter))]
        [ValidateRange(-360, 360)]
        public float FocusWordAngle { get; set; }

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
        /// Gets or sets which types of word rotations are used when drawing the word cloud.
        /// </summary>
        [Parameter()]
        [Alias()]
        public WordOrientations AllowRotation { get; set; } = WordOrientations.EitherVertical;

        /// <summary>
        /// Gets or sets the float value to scale the padding space around the words by.
        /// </summary>
        /// <value>The default value is 5.</value>
        [Parameter()]
        [Alias("Spacing")]
        public float Padding { get; set; } = 5;

        /// <summary>
        /// Get or sets the shape of backdrop to place behind each word.
        /// The default is no bubble.
        /// Be aware that circle or square bubbles will take up a lot more space than most words typically do;
        /// you may need to reduce the `-WordSize` parameter accordingly if you start getting warnings about words
        /// being skipped due to insufficient space.

        /// </summary>
        [Parameter()]
        public WordBubbleShape WordBubble { get; set; } = WordBubbleShape.None;

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
        /// Gets or sets whether to draw the cloud in monochrome (greyscale).
        /// </summary>
        [Parameter()]
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

        private float PaddingMultiplier { get => Padding * PADDING_BASE_SCALE; }

        private List<Task<List<string>>> _wordProcessingTasks;

        private int _colorSetIndex = 0;

        private int _progressId;

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

            _progressId = RandomInt();

            if (Monochrome.IsPresent)
            {
                ColorSet.TransformElements(c => c.AsMonochrome());
            }

            Shuffle(ColorSet);
        }

        /// <summary>
        /// Implements the ProcessRecord method for PSWordCloud.
        /// Spins up a Task&lt;IEnumerable&lt;string&gt;&gt; for each input text string to split them all
        /// asynchronously.
        /// </summary>
        protected override void ProcessRecord()
        {
            switch (ParameterSetName)
            {
                case FILE_SET:
                case FILE_FOCUS_SET:
                case COLOR_BG_SET:
                case COLOR_BG_FOCUS_SET:
                    List<string> text = NormalizeInput(InputObject);
                    _wordProcessingTasks ??= new List<Task<List<string>>>(GetEstimatedCapacity(InputObject));

                    foreach (var line in text)
                    {
                        var shortLine = Regex.Split(line, @"\r?\n")[0];
                        shortLine = shortLine.Length <= 32 ? shortLine : shortLine.Substring(0, 31) + ELLIPSIS;
                        WriteDebug($"Processing input text: {shortLine}");
                        _wordProcessingTasks.Add(ProcessInputAsync(line, IncludeWord, ExcludeWord));
                    }

                    break;
                default:
                    return;
            }
        }

        /// <summary>
        /// Implements the EndProcessing method for New-WordCloud.
        /// The majority of the word cloud drawing occurs here.
        /// </summary>
        protected override void EndProcessing()
        {
            if ((WordSizes == null || WordSizes.Count == 0)
                && (_wordProcessingTasks == null || _wordProcessingTasks.Count == 0))
            {
                // No input was supplied; exit stage left.
                return;
            }

            int currentWordNumber = 0;

            SKPath bubblePath = null;
            SKPath wordPath = null;
            SKRect viewbox = SKRect.Empty;
            SKBitmap backgroundImage = null;
            List<string> sortedWordList;

            Dictionary<string, float> wordScaleDictionary;
            switch (ParameterSetName)
            {
                case FILE_SET:
                case FILE_FOCUS_SET:
                case COLOR_BG_SET:
                case COLOR_BG_FOCUS_SET:
                    wordScaleDictionary = CancellableCollateWords();
                    break;
                case FILE_TABLE_SET:
                case FILE_FOCUS_TABLE_SET:
                case COLOR_BG_TABLE_SET:
                case COLOR_BG_FOCUS_TABLE_SET:
                    wordScaleDictionary = NormalizeWordScaleDictionary(WordSizes);
                    break;
                default:
                    throw new NotImplementedException("This parameter set has not defined an input handling method.");
            }

            // All words counted and in the dictionary.
            float highestWordFreq = wordScaleDictionary.Values.Max();

            if (MyInvocation.BoundParameters.ContainsKey(nameof(FocusWord)))
            {
                WriteDebug($"Adding focus word '{FocusWord}' to the dictionary.");
                wordScaleDictionary[FocusWord] = highestWordFreq *= FOCUS_WORD_SCALE;
            }

            // Get a sorted list of words by their sizes
            sortedWordList = new List<string>(SortWordList(wordScaleDictionary, MaxRenderedWords));

            try
            {
                if (MyInvocation.BoundParameters.ContainsKey(nameof(BackgroundImage)))
                {
                    // Set image size from the background size
                    WriteDebug($"Importing background image from '{_backgroundFullPath}'.");
                    backgroundImage = SKBitmap.Decode(_backgroundFullPath);
                    viewbox = new SKRectI(0, 0, backgroundImage.Width, backgroundImage.Height);
                }
                else
                {
                    // Set image size from default or specified size
                    viewbox = new SKRectI(0, 0, ImageSize.Width, ImageSize.Height);
                }

                using var clipRegion = GetClipRegion(viewbox, AllowOverflow.IsPresent);
                wordPath = new SKPath();

                float baseFontScale = GetBaseFontScale(
                    clipRegion.Bounds,
                    WordScale,
                    wordScaleDictionary.Values.Average(),
                    sortedWordList.Count,
                    Typeface);

                float maxWordWidth = AllowRotation == WordOrientations.None
                    ? viewbox.Width * MAX_WORD_WIDTH_PERCENT
                    : Math.Max(viewbox.Width, viewbox.Height) * MAX_WORD_WIDTH_PERCENT;

                using SKPaint brush = new SKPaint
                {
                    Typeface = Typeface
                };

                SKRect rect = SKRect.Empty;

                // Pre-test and adjust global scale based on the largest word.
                baseFontScale = GetAdjustedFontScale(
                    wordScaleDictionary,
                    sortedWordList[0],
                    maxWordWidth,
                    viewbox.Width * viewbox.Height,
                    baseFontScale);

                // Apply manual scaling from the user
                baseFontScale *= WordScale;
                WriteDebug($"Global font scale: {baseFontScale}");

                Dictionary<string, float> scaledWordSizes = GetScaledWordSizes(
                    sortedWordList,
                    wordScaleDictionary,
                    maxWordWidth,
                    viewbox.Width * viewbox.Height,
                    ref baseFontScale);

                // Remove all words that were cut from the final rendering list
                sortedWordList.RemoveAll(x => !scaledWordSizes.ContainsKey(x));

                using SKDynamicMemoryWStream outputStream = new SKDynamicMemoryWStream();
                using SKXmlStreamWriter xmlWriter = new SKXmlStreamWriter(outputStream);
                using SKCanvas canvas = SKSvgCanvas.Create(viewbox, xmlWriter);
                using SKRegion occupiedSpace = new SKRegion();

                brush.IsAutohinted = true;
                brush.IsAntialias = true;
                brush.Typeface = Typeface;

                if (ParameterSetName.StartsWith(FILE_SET))
                {
                    canvas.DrawBitmap(backgroundImage, 0, 0);
                }
                else if (BackgroundColor != SKColor.Empty)
                {
                    canvas.Clear(BackgroundColor);
                }

                SKPoint targetPoint;

                foreach (string word in sortedWordList.OrderByDescending(x => scaledWordSizes[x]))
                {
                    currentWordNumber++;

                    WriteDebug($"Scanning for draw location for '{word}'.");

                    SKColor wordColor;
                    SKColor bubbleColor = SKColor.Empty;

                    targetPoint = FindDrawLocation(
                        word,
                        brush,
                        scaledWordSizes[word],
                        currentWordNumber,
                        scaledWordSizes.Count,
                        viewbox,
                        clipRegion,
                        occupiedSpace,
                        out wordPath,
                        out bubblePath);

                    if (targetPoint != SKPoint.Empty)
                    {
                        if (WordBubble == WordBubbleShape.None)
                        {
                            wordColor = GetContrastingColor(BackgroundColor);
                        }
                        else
                        {
                            bubbleColor = GetContrastingColor(BackgroundColor);
                            wordColor = GetContrastingColor(bubbleColor);
                        }

                        WriteDebug($"Drawing '{word}' at [{targetPoint.X}, {targetPoint.Y}].");

                        wordPath.FillType = SKPathFillType.EvenOdd;

                        if (WordBubble != WordBubbleShape.None)
                        {
                            // If we're using word bubbles, the bubbles should more or less enclose the words.
                            occupiedSpace.CombineWithPath(bubblePath, SKRegionOperation.Union);

                            brush.IsStroke = false;
                            brush.Style = SKPaintStyle.Fill;
                            brush.Color = bubbleColor;
                            canvas.DrawPath(bubblePath, brush);
                        }
                        else
                        {
                            // If we're not using bubbles, record the exact space the word occupies.
                            occupiedSpace.CombineWithPath(wordPath, SKRegionOperation.Union);
                        }

                        brush.IsStroke = false;
                        brush.Style = SKPaintStyle.Fill;
                        brush.Color = wordColor;
                        canvas.DrawPath(wordPath, brush);

                        if (MyInvocation.BoundParameters.ContainsKey(nameof(StrokeWidth)))
                        {
                            brush.IsStroke = true;
                            brush.Style = SKPaintStyle.Stroke;
                            brush.Color = StrokeColor;

                            canvas.DrawPath(wordPath, brush);
                        }
                    }
                    else
                    {
                        WriteWarning($"Unable to find a place to draw '{word}'; skipping to next word.");
                    }
                }

                WriteDebug("Saving canvas data.");
                canvas.Flush();
                canvas.Dispose();
                outputStream.Flush();

                SaveSvgData(outputStream, viewbox);

                if (PassThru.IsPresent)
                {
                    WriteObject(InvokeProvider.Item.Get(Path), true);
                }
            }
            catch (Exception e)
            {
                WriteError(new ErrorRecord(e, "PSWordCloudError", ErrorCategory.InvalidResult, targetObject: null));
            }
            finally
            {
                WriteDebug("Disposing SkiaSharp objects.");

                wordPath?.Dispose();
                bubblePath?.Dispose();
                backgroundImage?.Dispose();

                // Write 'Completed' progress record
                WriteProgress(new ProgressRecord(_progressId, string.Empty, string.Empty)
                {
                    RecordType = ProgressRecordType.Completed
                });

                WriteProgress(new ProgressRecord(_progressId + 1, string.Empty, string.Empty)
                {
                    RecordType = ProgressRecordType.Completed
                });
            }
        }

        /// <summary>
        /// StopProcessing implementation for New-WordCloud.
        /// </summary>
        protected override void StopProcessing()
        {
            _cancel.Cancel();
        }

        #region HelperMethods

        /// <summary>
        /// Converts the input <paramref name="dictionary"/> into a usable form. Keys are all converted to string, and
        /// values are all converted to float.
        /// </summary>
        /// <param name="dictionary">The input dictionary to normalize.</param>
        /// <returns>The normalized Dictionary&lt;string, float&gt;.</returns>
        private Dictionary<string, float> NormalizeWordScaleDictionary(IDictionary dictionary)
        {
            var result = new Dictionary<string, float>(StringComparer.OrdinalIgnoreCase);

            WriteDebug("Processing -WordSizes input.");
            foreach (var word in dictionary.Keys)
            {
                try
                {
                    result.Add(
                        word.ConvertTo<string>(),
                        WordSizes[word].ConvertTo<float>());
                }
                catch (Exception e)
                {
                    WriteWarning($"Skipping entry '{word}' due to error converting key or value: {e.Message}.");
                    WriteDebug($"Entry type: key - {word.GetType().FullName} ; value - {WordSizes[word].GetType().FullName}");
                }
            }

            return result;
        }

        /// <summary>
        /// Waits for all the word processing tasks to complete and then counts everything before building the word
        /// frequency dictionary.
        /// If StopProcessing() is called during the tasks' operation
        /// </summary>
        /// <returns></returns>
        private Dictionary<string, float> CancellableCollateWords()
        {
            var dictionary = new Dictionary<string, float>(StringComparer.OrdinalIgnoreCase);
            WriteDebug("Waiting for word processing tasks to finish.");
            var processingTasks = Task.WhenAll(_wordProcessingTasks);

            var waitHandles = new[] { ((IAsyncResult)processingTasks).AsyncWaitHandle, _cancel.Token.WaitHandle };
            if (WaitHandle.WaitAny(waitHandles) == 1)
            {
                throw new PipelineStoppedException();
            }

            WriteDebug("Word processing tasks complete.");
            var result = processingTasks.GetAwaiter().GetResult();

            WriteDebug("Counting words and populating scaling dictionary.");
            foreach (var lineWords in result)
            {
                CountWords(lineWords, dictionary);
            }

            return dictionary;
        }

        private static SKRegion GetClipRegion(SKRect viewbox, bool allowOverflow)
        {
            var clipRegion = new SKRegion();
            if (allowOverflow)
            {
                clipRegion.SetRect(
                    SKRectI.Round(SKRect.Inflate(
                        viewbox,
                        viewbox.Width * (BLEED_AREA_SCALE - 1),
                        viewbox.Height * (BLEED_AREA_SCALE - 1))));
            }
            else
            {
                clipRegion.SetRect(SKRectI.Round(viewbox));
            }

            return clipRegion;
        }

        private Dictionary<string, float> GetScaledWordSizes(
            IList<string> sortedWordList,
            IDictionary<string, float> wordScaleDictionary,
            float maxWordWidth,
            float imageArea,
            ref float baseFontScale)
        {
            var scaledWordSizes = new Dictionary<string, float>(
                sortedWordList.Count,
                StringComparer.OrdinalIgnoreCase);
            using var brush = new SKPaint();

            foreach (string word in sortedWordList)
            {
                float adjustedWordScale = ScaleWordSize(
                    wordScaleDictionary[word],
                    baseFontScale,
                    wordScaleDictionary);

                brush.Prepare(adjustedWordScale, StrokeWidth);

                var textRect = brush.GetTextPath(word, 0, 0).ComputeTightBounds();
                var adjustedTextWidth = textRect.Width * (1 + PaddingMultiplier) + StrokeWidth * 2 * STROKE_BASE_SCALE;

                if (!AllowOverflow.IsPresent
                    && (adjustedTextWidth > maxWordWidth
                        || textRect.Width * textRect.Height > imageArea * MAX_WORD_AREA_PERCENT))
                {
                    baseFontScale *= 0.95f;
                    return GetScaledWordSizes(sortedWordList, wordScaleDictionary, maxWordWidth, imageArea, ref baseFontScale);
                }

                scaledWordSizes[word] = adjustedWordScale;
            }

            return scaledWordSizes;
        }

        private float GetAdjustedFontScale(
            IDictionary<string, float> wordFrequencyTable,
            string largestWord,
            float maxWordWidth,
            float imageArea,
            float baseFontScale)
        {
            using var brush = new SKPaint();

            float adjustedWordSize = ScaleWordSize(
                wordFrequencyTable[largestWord],
                baseFontScale,
                wordFrequencyTable);

            brush.Prepare(adjustedWordSize, StrokeWidth);

            var textRect = brush.GetTextPath(largestWord, 0, 0).ComputeTightBounds();
            var adjustedTextWidth = textRect.Width * (1 + PaddingMultiplier) + StrokeWidth * 2 * STROKE_BASE_SCALE;

            if (adjustedTextWidth > maxWordWidth
                || textRect.Width * textRect.Height < imageArea * MAX_WORD_AREA_PERCENT)
            {
                baseFontScale *= 1.05f;
                return GetAdjustedFontScale(wordFrequencyTable, largestWord, maxWordWidth, imageArea, baseFontScale);
            }

            return baseFontScale;
        }

        /// <summary>
        /// Scans the image space to find an available draw location for the word, taking into account the already-drawn
        /// words in <paramref name="occupiedSpace"/> and avoiding collision.
        /// </summary>
        /// <param name="word">The word being drawn.</param>
        /// <param name="brush"></param>
        /// <param name="wordSize"></param>
        /// <param name="currentWord"></param>
        /// <param name="totalWords"></param>
        /// <param name="viewbox"></param>
        /// <param name="clipRegion"></param>
        /// <param name="occupiedSpace"></param>
        /// <param name="wordPath"></param>
        /// <param name="bubblePath"></param>
        /// <returns></returns>
        private SKPoint FindDrawLocation(
            string word,
            SKPaint brush,
            float wordSize,
            int currentWord,
            int totalWords,
            SKRect viewbox,
            SKRegion clipRegion,
            SKRegion occupiedSpace,
            out SKPath wordPath,
            out SKPath bubblePath)
        {
            var wordProgress = new ProgressRecord(
                _progressId,
                "Drawing word cloud...",
                $"Finding space for word: '{word}'...");

            var pointProgress = new ProgressRecord(
                _progressId + 1,
                "Scanning available space...",
                "Scanning radial points...")
            {
                ParentActivityId = _progressId
            };

            float[] availableAngles = currentWord == 1
                && MyInvocation.BoundParameters.ContainsKey(nameof(FocusWordAngle))
                    ? new[] { FocusWordAngle }
                    : GetDrawAngles(AllowRotation);

            var centrePoint = new SKPoint(viewbox.MidX, viewbox.MidY);
            float aspectRatio = viewbox.Width / viewbox.Height;
            float inflationValue = 2 * wordSize * (PaddingMultiplier + StrokeWidth * STROKE_BASE_SCALE);

            // Max radius should reach to the corner of the image; location is top-left of the box
            float maxRadius = SKPoint.Distance(viewbox.Location, centrePoint);

            if (AllowOverflow)
            {
                maxRadius *= BLEED_AREA_SCALE;
            }

            wordPath = null;
            bubblePath = null;
            SKRect wordBounds;
            brush.Prepare(wordSize, StrokeWidth);

            var percentComplete = 100f * currentWord / totalWords;
            wordProgress.StatusDescription = string.Format(
                "Draw: \"{0}\" [Size: {1:0}] ({2} of {3})",
                word,
                brush.TextSize,
                currentWord,
                totalWords);
            wordProgress.PercentComplete = (int)Math.Round(percentComplete);
            WriteProgress(wordProgress);

            foreach (var drawAngle in availableAngles)
            {
                wordPath = brush.GetTextPath(word, 0, 0);
                wordBounds = wordPath.TightBounds;

                SKMatrix rotation = SKMatrix.MakeRotationDegrees(drawAngle, wordBounds.MidX, wordBounds.MidY);
                wordPath.Transform(rotation);

                for (
                    float radius = 0;
                    radius <= maxRadius;
                    radius += GetRadiusIncrement(
                        wordSize,
                        DistanceStep,
                        maxRadius,
                        inflationValue,
                        percentComplete))
                {
                    var radialPoints = GetOvalPoints(centrePoint, radius, RadialStep, aspectRatio);
                    var totalPoints = radialPoints.Count;
                    var pointsChecked = 0;

                    foreach (var point in radialPoints)
                    {
                        pointsChecked++;
                        if (!clipRegion.Contains(point) && point != centrePoint)
                        {
                            continue;
                        }

                        pointProgress.Activity = string.Format(
                            "Finding available space to draw at angle: {0}",
                            drawAngle);
                        pointProgress.StatusDescription = string.Format(
                            "Checking [Point:{0,8:N2}, {1,8:N2}] ({2,4} / {3,4}) at [Radius: {4,8:N2}]",
                            point.X,
                            point.Y,
                            pointsChecked,
                            totalPoints,
                            radius);

                        WriteProgress(pointProgress);

                        var pathMidpoint = new SKPoint(wordBounds.MidX, wordBounds.MidY);

                        wordPath.Offset(point - pathMidpoint);
                        wordBounds = wordPath.TightBounds;

                        wordBounds.Inflate(inflationValue, inflationValue);

                        if (WordBubble == WordBubbleShape.None)
                        {
                            if (currentWord == 1)
                            {
                                // First word always gets drawn, and will be in the centre.
                                return wordBounds.Location;
                            }

                            if (wordBounds.FallsOutside(clipRegion))
                            {
                                continue;
                            }

                            if (!occupiedSpace.IntersectsRect(wordBounds))
                            {
                                return wordBounds.Location;
                            }
                        }
                        else
                        {
                            bubblePath = new SKPath();

                            SKRoundRect wordBubble;
                            float bubbleRadius;

                            switch (WordBubble)
                            {
                                case WordBubbleShape.Rectangle:
                                    bubbleRadius = wordBounds.Height / 16;
                                    wordBubble = new SKRoundRect(wordBounds, bubbleRadius, bubbleRadius);
                                    bubblePath.AddRoundRect(wordBubble);
                                    break;

                                case WordBubbleShape.Square:
                                    bubbleRadius = Math.Max(wordBounds.Width, wordBounds.Height) / 16;
                                    wordBubble = new SKRoundRect(wordBounds.GetEnclosingSquare(), bubbleRadius, bubbleRadius);
                                    bubblePath.AddRoundRect(wordBubble);
                                    break;

                                case WordBubbleShape.Circle:
                                    bubbleRadius = Math.Max(wordBounds.Width, wordBounds.Height) / 2;
                                    bubblePath.AddCircle(wordBounds.MidX, wordBounds.MidY, bubbleRadius);
                                    break;

                                case WordBubbleShape.Oval:
                                    bubblePath.AddOval(wordBounds);
                                    break;
                            }

                            if (currentWord == 1)
                            {
                                // First word always gets drawn, and will be in the centre.
                                return wordBounds.Location;
                            }

                            if (wordBounds.FallsOutside(clipRegion))
                            {
                                continue;
                            }

                            if (!occupiedSpace.IntersectsPath(bubblePath))
                            {
                                return wordBounds.Location;
                            }
                        }

                        if (point == centrePoint)
                        {
                            // No point checking more than a single point at the origin
                            break;
                        }
                    }
                }
            }

            return SKPoint.Empty;
        }

        /// <summary>
        /// Gets the next available color from the current set.
        /// If the set's end is reached, it will loop back to the beginning of the set again.
        /// </summary>
        private SKColor GetNextColor()
        {
            if (_colorSetIndex >= ColorSet.Length)
            {
                _colorSetIndex = 0;
            }

            return ColorSet[_colorSetIndex++];
        }

        /// <summary>
        /// Gets the next available color that is sufficiently visually distinct from the reference color.
        /// </summary>
        /// <param name="reference">A color that should contrast with the returned color.</param>
        private SKColor GetContrastingColor(SKColor reference)
        {
            SKColor result;
            do
            {
                result = GetNextColor();
            }
            while (!result.IsDistinctFrom(reference));

            return result;
        }

        /// <summary>
        /// Returns a shuffled set of possible angles determined by the WordOrientations value provided.
        /// </summary>
        private static float[] GetDrawAngles(WordOrientations permittedRotations)
        {
            return permittedRotations switch
            {
                WordOrientations.Vertical => Shuffle(new float[] { 0, 90 }),
                WordOrientations.FlippedVertical => Shuffle(new float[] { 0, -90 }),
                WordOrientations.EitherVertical => Shuffle(new float[] { 0, 90, -90 }),
                WordOrientations.UprightDiagonals => Shuffle(new float[] { 0, -90, -45, 45, 90 }),
                WordOrientations.InvertedDiagonals => Shuffle(new float[] { 90, 135, -135, -90, 180 }),
                WordOrientations.AllDiagonals => Shuffle(new float[] { 45, 90, 135, 180, -135, -90, -45, 0 }),
                WordOrientations.AllUpright => RandomAngles(-90, 91),
                WordOrientations.AllInverted => RandomAngles(90, 271),
                WordOrientations.All => RandomAngles(0, 361),
                _ => new float[] { 0 },
            };
        }

        /// <summary>
        /// Yields a set of random angles between <paramref name="minAngle"/> and <paramref name="maxAngle"/>.
        /// </summary>
        private static float[] RandomAngles(int minAngle, int maxAngle)
        {
            var angles = new float[RandomInt(4, 12)];
            for (var index = 0; index < angles.Length; index++)
            {
                angles[index] = RandomFloat(minAngle, maxAngle);
            }

            return angles;
        }

        /// <summary>
        /// Save the written SVG data to the provided PSProvider path.
        /// Since SkiaSharp does not write a viewbox attribute into the SVG, this method handles that as well.
        /// </summary>
        /// <param name="outputStream">The memory stream containing the SVG data.</param>
        /// <param name="viewbox">The visible area of the image.</param>
        private void SaveSvgData(SKDynamicMemoryWStream outputStream, SKRect viewbox)
        {
            string[] path = new[] { Path };

            if (InvokeProvider.Item.Exists(Path, force: true, literalPath: true))
            {
                WriteDebug($"Clearing existing content from '{Path}'.");
                try
                {
                    InvokeProvider.Content.Clear(path, force: false, literalPath: true);
                }
                catch (Exception e)
                {
                    // Unconditionally suppress errors from the Content.Clear() operation. Errors here may indicate that
                    // a provider is being written to that does not support the Content.Clear() interface, or that there
                    // is no existing item to clear.
                    // In either case, an error here does not necessarily mean we cannot write the data, so we can
                    // ignore this error. If there is an access denied error, it will be more clear to the user if we
                    // surface that from the Content.Write() interface in any case.
                    WriteDebug($"Error encountered while clearing content for item '{path}'. {e.Message}");
                }
            }

            using SKData data = outputStream.DetachAsData();
            using var reader = new StreamReader(data.AsStream());
            using var writer = InvokeProvider.Content.GetWriter(path, force: false, literalPath: true).First();

            var imageXml = new XmlDocument();
            imageXml.LoadXml(reader.ReadToEnd());

            var svgElement = imageXml.GetElementsByTagName("svg")[0] as XmlElement;
            if (svgElement.GetAttribute("viewbox") == string.Empty)
            {
                svgElement.SetAttribute(
                    "viewbox",
                    $"{viewbox.Location.X} {viewbox.Location.Y} {viewbox.Width} {viewbox.Height}");
            }

            WriteDebug($"Saving data to '{Path}'.");
            writer.Write(new[] { imageXml.GetPrettyString() });
            writer.Close();
        }

        /// <summary>
        /// Check the type of the input object and return a probable count so we can reasonably estimate necessary
        /// capacity for processing.
        /// </summary>
        /// <param name="inputObject"></param>
        /// <returns></returns>
        private int GetEstimatedCapacity(PSObject inputObject) => inputObject.BaseObject switch
        {
            string _ => 1,
            IList list => list.Count,
            _ => 8
        };

        /// <summary>
        /// Process a given input object and convert it to a string (or multiple strings, if there are more than one).
        /// </summary>
        /// <param name="input">One or more input objects.</param>
        private List<string> NormalizeInput(PSObject input)
        {
            var list = new List<string>();
            string value;
            switch (input.BaseObject)
            {
                case string s2:
                    list.Add(s2);
                    break;

                case string[] sa:
                    list.AddRange(sa);
                    break;

                default:
                    IEnumerable enumerable;
                    try
                    {
                        enumerable = LanguagePrimitives.GetEnumerable(input.BaseObject);
                    }
                    catch
                    {
                        break;
                    }

                    if (enumerable != null)
                    {
                        foreach (var item in enumerable)
                        {
                            try
                            {
                                value = item.ConvertTo<string>();
                            }
                            catch
                            {
                                break;
                            }

                            list.Add(value);
                        }

                        break;
                    }

                    try
                    {
                        value = input.ConvertTo<string>();
                    }
                    catch
                    {
                        break;
                    }

                    list.Add(value);
                    break;
            }

            return list;
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
        private static float GetBaseFontScale(
            SKRect space,
            float baseScale,
            float averageWordFrequency,
            int wordCount,
            SKTypeface typeface)
        {
            var FontScale = WCUtils.GetFontScale(typeface);
            return baseScale * FontScale * Math.Max(space.Height, space.Width) / (averageWordFrequency * wordCount);
        }

        /// <summary>
        /// Scale each word by the base word scale value and determine its final font size.
        /// </summary>
        /// <param name="baseSize">The base size for the font.</param>
        /// <param name="globalScale">The global scaling factor.</param>
        /// <param name="scaleDictionary">The dictionary of word scales containing their base sizes.</param>
        /// <returns></returns>
        private static float ScaleWordSize(
            float baseSize,
            float globalScale,
            IDictionary<string, float> scaleDictionary)
        {
            return baseSize / scaleDictionary.Values.Max() * globalScale * (1 + RandomFloat() / 5);
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
            float wordSize,
            float distanceStep,
            float maxRadius,
            float padding,
            float percentComplete)
        {
            var wordScaleFactor = (1 + padding) * wordSize / 360;
            var stepScale = distanceStep / 15;
            var minRadiusIncrement = percentComplete / 100 * maxRadius / 10;

            var radiusIncrement = minRadiusIncrement * stepScale + wordScaleFactor;

            return radiusIncrement;
        }

        /// <summary>
        /// Scans in an ovoid pattern at a given radius to get a set of points to check for sufficient drawing space.
        /// </summary>
        /// <param name="centre">The centre point of the image.</param>
        /// <param name="radius">The current radius we're scanning at.</param>
        /// <param name="radialStep">The current radial stepping value.</param>
        /// <param name="aspectRatio">The aspect ratio of the canvas.</param>
        /// <returns></returns>
        private static List<SKPoint> GetOvalPoints(
            SKPoint centre,
            float radius,
            float radialStep,
            float aspectRatio = 1)
        {
            var result = new List<SKPoint>();
            if (radius == 0)
            {
                result.Add(centre);
                return result;
            }

            Complex point;

            var baseRadialPoints = 7;
            var baseAngleIncrement = 360 / baseRadialPoints;
            float angleIncrement = baseAngleIncrement / (float)Math.Sqrt(radius);

            bool clockwise = RandomFloat() > 0.5;

            float angle = RandomInt(0, 4) switch
            {
                1 => 90,
                2 => 180,
                3 => 270,
                _ => 0
            };

            float maxAngle;
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
                result.Add(new SKPoint(centre.X + (float)point.Real * aspectRatio, centre.Y + (float)point.Imaginary));

                angle += angleIncrement * (radialStep / 15);
            } while (clockwise ? angle <= maxAngle : angle >= maxAngle);

            return result;
        }

        /// <summary>
        /// Retrieves a random floating-point number between 0 and 1.
        /// </summary>
        private static float RandomFloat()
        {
            lock (_randomLock)
            {
                return (float)Random.NextDouble();
            }
        }

        /// <summary>
        /// Retrieves a random floating-point number between <paramref name="min"/> and <paramref name="max"/>.
        /// </summary>
        private static float RandomFloat(float min, float max)
        {
            if (min > max)
            {
                return max;
            }

            lock (_randomLock)
            {
                var range = max - min;
                return (float)Random.NextDouble() * range + min;
            }
        }

        /// <summary>
        /// Retrieves a random int value.
        /// </summary>
        private static int RandomInt()
        {
            lock (_randomLock)
            {
                return Random.Next();
            }
        }

        /// <summary>
        /// Retrieves a random int value within the specified bounds.
        /// </summary>
        /// <param name="min">The minimum bound.</param>
        /// <param name="max">The maximum bound.</param>
        private static int RandomInt(int min, int max)
        {
            lock (_randomLock)
            {
                return Random.Next(min, max);
            }
        }

        /// <summary>
        /// Performs an in-place shuffle of the input array, randomly shuffling its contents.
        /// </summary>
        /// <param name="items">The array of items to be shuffled.</param>
        /// <typeparam name="T">The type of the array.</typeparam>
        private static IList<T> Shuffle<T>(IList<T> items)
        {
            lock (_randomLock)
            {
                return Random.Shuffle(items);
            }
        }

        /// <summary>
        /// Asynchronous method used to quickly process large amounts of text input into words.
        /// </summary>
        /// <param name="line">The text to split and process.</param>
        /// <returns>An enumerable string collection of all words in the input, with stopwords stripped out.</returns>
        private async Task<List<string>> ProcessInputAsync(
            string line,
            string[] includeWords = null,
            string[] excludeWords = null)
        {
            return await Task.Run(() => TrimAndSplitWords(line)
                .Where(x => SelectWord(x, includeWords, excludeWords, AllowStopWords.IsPresent))
                .ToList());
        }

        /// <summary>
        /// Enumerates and returns each word from the input text.
        /// </summary>
        /// <param name="text">A string of text to extract words from.</param>
        private IEnumerable<string> TrimAndSplitWords(string text)
        {
            foreach (var word in text.Split(_splitChars, StringSplitOptions.RemoveEmptyEntries))
            {
                yield return Regex.Replace(word, @"^[^a-zA-Z0-9]+|[^a-zA-Z0-9]+$", string.Empty);
            }
        }

        /// <summary>
        /// Determines whether a given word is usable for the word cloud, taking into account stopwords and
        /// user selected include or exclude word lists.
        /// </summary>
        /// <param name="word">The word in question.</param>
        /// <param name="includeWords">A reference list of desired words, overridingthe stopwords or exclude list.</param>
        /// <param name="excludeWords">A reference list of undesired words, effectively impromptu stopwords.</param>
        private static bool SelectWord(string word, string[] includeWords, string[] excludeWords, bool allowStopWords)
        {
            if (includeWords?.Contains(word, StringComparer.OrdinalIgnoreCase) == true)
            {
                return true;
            }

            if (excludeWords?.Contains(word, StringComparer.OrdinalIgnoreCase) == true)
            {
                return false;
            }

            if (!allowStopWords && _stopWords.Contains(word, StringComparer.OrdinalIgnoreCase))
            {
                return false;
            }

            if (Regex.Replace(word, "[^a-zA-Z-]", string.Empty).Length < 2)
            {
                return false;
            }

            return true;
        }

        #endregion HelperMethods
    }
}
