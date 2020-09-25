using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Provider;
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
    /// </summary>
    [Cmdlet(VerbsCommon.New, "WordCloud", DefaultParameterSetName = COLOR_BG_SET,
        HelpUri = "https://github.com/vexx32/PSWordCloud/blob/main/docs/New-WordCloud.md")]
    [Alias("wordcloud", "nwc", "wcloud")]
    [OutputType(typeof(FileInfo))]
    public class NewWordCloudCommand : PSCmdlet
    {

        #region Constants

        internal const string COLOR_BG_SET = "ColorBackground";
        internal const string COLOR_BG_FOCUS_SET = "ColorBackground-FocusWord";
        internal const string COLOR_BG_FOCUS_TABLE_SET = "ColorBackground-FocusWord-WordTable";
        internal const string COLOR_BG_TABLE_SET = "ColorBackground-WordTable";
        internal const string FILE_SET = "FileBackground";
        internal const string FILE_FOCUS_SET = "FileBackground-FocusWord";
        internal const string FILE_FOCUS_TABLE_SET = "FileBackground-FocusWord-WordTable";
        internal const string FILE_TABLE_SET = "FileBackground-WordTable";

        #endregion

        #region Static Members

        private static LockingRandom? _lockingRandom;
        internal static LockingRandom SafeRandom => _lockingRandom ??= new LockingRandom();

        private readonly CancellationTokenSource _cancel = new CancellationTokenSource();

        #endregion

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
        [Alias("InputString", "Text", "String", "Document", "Page")]
        [AllowEmptyString()]
        public PSObject? InputObject { get; set; }

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
        public IDictionary? WordSizes { get; set; }

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
        public string Path { get; set; } = string.Empty;

        private string? _backgroundFullPath;
        /// <summary>
        /// Gets or sets the path to the background image to be used as a base for the final word cloud image.
        /// </summary>
        [Parameter(Mandatory = true, ParameterSetName = FILE_SET)]
        [Parameter(Mandatory = true, ParameterSetName = FILE_FOCUS_SET)]
        [Parameter(Mandatory = true, ParameterSetName = FILE_FOCUS_TABLE_SET)]
        [Parameter(Mandatory = true, ParameterSetName = FILE_TABLE_SET)]
        public string? BackgroundImage
        {
            get => _backgroundFullPath!;
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
        public string? FocusWord { get; set; }

        /// <summary>
        /// Gets or sets the angle to draw the focus word at in degrees. The default is
        /// </summary>
        /// <value></value>
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
        public string[]? ExcludeWord { get; set; }

        /// <summary>
        /// <para>Gets or sets the words to be explicitly included in rendering of the cloud.</para>
        /// <para>This can be used to override specific words normally excluded by the StopWords list.</para>
        /// </summary>
        /// <value></value>
        [Parameter()]
        [Alias()]
        public string[]? IncludeWord { get; set; }

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
        [Parameter(ParameterSetName = COLOR_BG_SET)]
        [Parameter(ParameterSetName = COLOR_BG_FOCUS_SET)]
        [Parameter(ParameterSetName = FILE_SET)]
        [Parameter(ParameterSetName = FILE_FOCUS_SET)]
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
        /// Gets or sets the value that determines whether or not to retrieve and output the item that
        /// represents the completed word cloud when processing is completed.
        /// </summary>
        [Parameter()]
        public SwitchParameter PassThru { get; set; }

        #endregion

        #region Private Variables

        private float PaddingMultiplier { get => Padding * Constants.PaddingBaseScale; }

        private readonly ConcurrentBag<string> _processedWords = new ConcurrentBag<string>();
        private readonly List<EventWaitHandle> _waitHandles = new List<EventWaitHandle>();

        private int _colorSetIndex = 0;

        private int _progressId;

        #endregion

        #region Overrides

        /// <summary>
        /// Implements the <see cref="Cmdlet.BeginProcessing"/> method for <see cref="NewWordCloudCommand"/>.
        /// Instantiates the random number generator, and organises the base color set for the cloud.
        /// </summary>
        protected override void BeginProcessing()
        {
            InitializeSafeRandom();
            SetProgressId();
            PrepareColorSet();
        }

        /// <summary>
        /// Implements the <see cref="ProcessRecord"/> method for <see cref="NewWordCloudCommand"/>.
        /// </summary>
        protected override void ProcessRecord()
        {
            switch (ParameterSetName)
            {
                case FILE_SET:
                case FILE_FOCUS_SET:
                case COLOR_BG_SET:
                case COLOR_BG_FOCUS_SET:
                    if (InputObject is null)
                    {
                        return;
                    }

                    QueueInputProcessing(InputObject);
                    break;

                default:
                    return;
            }
        }

        /// <summary>
        /// Implements the <see cref="EndProcessing"/> method for <see cref="NewWordCloudCommand"/>.
        /// The majority of the word cloud drawing occurs here.
        /// </summary>
        protected override void EndProcessing()
        {
            bool wordSizesWereSpecified = WordSizes?.Count > 0;
            bool hasTextInput = _waitHandles.Count > 0;
            if (!(wordSizesWereSpecified || hasTextInput))
            {
                WriteDebug("No input was received. Ending processing.");
                return;
            }

            CreateWordCloud();
        }

        /// <summary>
        /// StopProcessing implementation for New-WordCloud.
        /// </summary>
        protected override void StopProcessing()
        {
            // Cancellation is registered in both the word-processing and word placement code regions,
            // which comprise the majority of the "slow" / "blocking" behaviour of the cmdlet.
            _cancel.Cancel();
        }

        #endregion

        private void CreateWordCloud()
        {
            using var image = CreateImage();
            BackgroundColor = image.BackgroundColor;

            IReadOnlyList<Word> wordList = GetScaledWords(image);

            try
            {
                DrawAllWordsOnCanvas(wordList, image);

                WriteDebug($"Saving canvas data to {string.Join(',', Path)}.");
                SaveSvgData(image, Path);

                if (PassThru.IsPresent)
                {
                    WriteObject(InvokeProvider.Item.Get(Path), enumerateCollection: true);
                }
            }
            catch (Exception e)
            {
                WriteError(new ErrorRecord(e, "PSWordCloudError", ErrorCategory.InvalidResult, targetObject: null));
            }
            finally
            {
                WriteProgressCompleted();
            }
        }

        #region Helpers - Setup/Preparation

        private void InitializeSafeRandom()
        {
            _lockingRandom = MyInvocation.BoundParameters.ContainsKey(nameof(RandomSeed))
                ? new LockingRandom(RandomSeed)
                : new LockingRandom();
        }

        private void SetProgressId()
        {
            _progressId = SafeRandom.GetRandomInt();
        }

        private void PrepareColorSet()
        {
            if (Monochrome.IsPresent)
            {
                ConvertColorSetToMonochrome();
            }

            SafeRandom.Shuffle(ColorSet);
            ColorSet = ColorSet.Take(MaxColors).ToArray();
        }

        private void ConvertColorSetToMonochrome() => ColorSet.TransformElements(c => c.AsMonochrome());

        #endregion

        #region Helpers - Processing Input

        private void QueueInputProcessing(PSObject inputObject)
        {
            var waitHandle = new EventWaitHandle(initialState: false, EventResetMode.ManualReset);
            ThreadPool.QueueUserWorkItem(StartProcessingInput, (inputObject, waitHandle), preferLocal: false);
            _waitHandles.Add(waitHandle);
        }

        private void StartProcessingInput((PSObject, EventWaitHandle) state)
        {
            (PSObject? inputObject, EventWaitHandle waitHandle) = state;
            ProcessInput(
                CreateStringList(PSObject.AsPSObject(inputObject)),
                IncludeWord,
                ExcludeWord);

            waitHandle.Set();
        }

        private IReadOnlyList<string> CreateStringList(PSObject input)
            => input.BaseObject switch
            {
                string s => new[] { s },
                string[] sa => sa,
                _ => GetStrings(input.BaseObject)
            };

        private IReadOnlyList<string> GetStrings(object baseObject)
        {
            IEnumerable enumerable = LanguagePrimitives.GetEnumerable(baseObject);

            if (enumerable is not null)
            {
                return GetStringsFromEnumerable(enumerable);
            }

            return GetStringOrEmptyList(baseObject);
        }

        private IReadOnlyList<string> GetStringOrEmptyList(object baseObject)
        {
            try
            {
                return new[] { baseObject.ConvertTo<string>() };
            }
            catch
            {
                return Array.Empty<string>();
            }
        }

        private IReadOnlyList<string> GetStringsFromEnumerable(IEnumerable enumerable)
        {
            var strings = new List<string>();
            foreach (var item in enumerable)
            {
                if (item is null)
                {
                    continue;
                }

                try
                {
                    strings.Add(item.ConvertTo<string>());
                }
                catch
                {
                    continue;
                }
            }

            return strings;
        }

        private void ProcessInput(
            IReadOnlyList<string> lines,
            IReadOnlyList<string>? includeWords = null,
            IReadOnlyList<string>? excludeWords = null)
        {
            var words = TrimAndSplitWords(lines);
            for (int index = 0; index < words.Count; index++)
            {
                if (KeepWord(words[index], includeWords, excludeWords, AllowStopWords.IsPresent))
                {
                    _processedWords.Add(words[index]);
                }
            }
        }

        private IReadOnlyList<string> TrimAndSplitWords(IReadOnlyList<string> text)
        {
            var wordList = new List<string>(text.Count);
            for (int i = 0; i < text.Count; i++)
            {
                string[] initialWords = text[i].Split(WCUtils.SplitChars, StringSplitOptions.RemoveEmptyEntries);

                for (int index = 0; index < initialWords.Length; index++)
                {
                    var word = Regex.Replace(initialWords[index], @"^[^a-zA-Z0-9]+|[^a-zA-Z0-9]+$", string.Empty);
                    wordList.Add(word);
                }
            }

            return wordList;
        }

        private static bool KeepWord(
            string word,
            IReadOnlyList<string>? includeWords,
            IReadOnlyList<string>? excludeWords,
            bool allowStopWords)
        {
            if (includeWords?.Contains(word, StringComparer.OrdinalIgnoreCase) == true)
            {
                return true;
            }

            if (excludeWords?.Contains(word, StringComparer.OrdinalIgnoreCase) == true)
            {
                return false;
            }

            if (!allowStopWords && WCUtils.IsStopWord(word))
            {
                return false;
            }

            string onlyLetters = Regex.Replace(word, "[^a-zA-Z-]", string.Empty);
            return onlyLetters.Length >= 2;
        }

        #endregion

        #region Helpers - Collating Input

        private IReadOnlyList<Word> GetScaledWords(Image image)
        {
            IReadOnlyList<Word> wordList = GetRelativeWordSizes(ParameterSetName);
            ThrowIfEmpty(wordList);

            float maxWordWidth = MyInvocation.BoundParameters.ContainsKey(nameof(FocusWord))
                ? GetMaxWordWidth(image.Viewbox, FocusWordAngle)
                : GetMaxWordWidth(image.Viewbox, AllowRotation);

            float scaleFactor = GetWordScaleFactor(wordList, maxWordWidth, image.ClippingBounds);

            WriteDebug($"Global font scale: {scaleFactor}");

            return ScaleWordSizes(wordList, maxWordWidth, scaleFactor, WordScale);
        }

        private IReadOnlyList<Word> GetRelativeWordSizes(string parameterSet)
        {
            Dictionary<string, float> wordScaleDictionary;
            int maxWords = 0;
            switch (parameterSet)
            {
                case FILE_SET:
                case FILE_FOCUS_SET:
                case COLOR_BG_SET:
                case COLOR_BG_FOCUS_SET:
                    wordScaleDictionary = GetWordCountDictionary();
                    maxWords = MaxRenderedWords;
                    break;
                case FILE_TABLE_SET:
                case FILE_FOCUS_TABLE_SET:
                case COLOR_BG_TABLE_SET:
                case COLOR_BG_FOCUS_TABLE_SET:
                    wordScaleDictionary = GetProcessedWordScaleDictionary(WordSizes!);
                    break;
                default:
                    throw new NotImplementedException("This parameter set has not defined an input handling method.");
            }

            return GetWordListFromDictionary(wordScaleDictionary, maxWords);
        }

        private float GetMaxWordWidth(SKRect viewbox, float focusWordAngle)
            => WCUtils.AngleIsMostlyVertical(focusWordAngle)
                ? viewbox.Height * Constants.MaxWordWidthPercent
                : viewbox.Width * Constants.MaxWordWidthPercent;

        private float GetMaxWordWidth(SKRect viewbox, WordOrientations permittedOrientations)
            => permittedOrientations == WordOrientations.None
                ? viewbox.Width * Constants.MaxWordWidthPercent
                : Math.Max(viewbox.Width, viewbox.Height) * Constants.MaxWordWidthPercent;

        private float GetWordScaleFactor(IReadOnlyList<Word> wordList, float maxWordWidth, SKRect drawableBounds)
        {
            float estimatedScale = EstimateWordScale(drawableBounds, wordList, Typeface);
            return ConstrainScaleByWordWidth(maxWordWidth, wordList[0], estimatedScale);
        }

        private float EstimateWordScale(
            SKRect canvasRect,
            IReadOnlyList<Word> wordList,
            SKTypeface typeface)
        {
            var stats = new WordListStatistics(wordList);

            float fontCharArea = WCUtils.GetAverageCharArea(typeface);
            float estimatedPadding = (float)Math.Sqrt(fontCharArea) * Constants.PaddingBaseScale / stats.AverageFrequency;
            float estimatedWordArea = fontCharArea * stats.AverageLength * stats.AverageFrequency * stats.Count + estimatedPadding;
            float canvasArea = canvasRect.Height * canvasRect.Width;

            return canvasArea * Constants.MaxWordAreaPercent / estimatedWordArea;
        }

        private float ConstrainScaleByWordWidth(float maxWordWidth, Word largestWord, float fontScale)
        {
            float largestWordSize = largestWord.Scale(fontScale);

            using SKPaint brush = WCUtils.GetBrush(largestWordSize, StrokeWidth * Constants.StrokeBaseScale, Typeface);
            largestWord.Path = brush.GetTextPath(largestWord.Text, x: 0, y: 0);

            float padding = GetPaddingValue(largestWord);
            float effectiveWidth = Math.Max(Constants.MinEffectiveWordWidth, largestWord.Text.Length);
            float adjustedWidth = largestWord.Path.Bounds.Width * (effectiveWidth / largestWord.Text.Length) + padding;
            if (adjustedWidth > maxWordWidth)
            {
                return ConstrainScaleByWordWidth(
                    maxWordWidth,
                    largestWord,
                    fontScale * Constants.MaxWordWidthPercent * (maxWordWidth / adjustedWidth));
            }

            return fontScale;
        }

        private Dictionary<string, float> GetWordCountDictionary()
        {
            CompleteInputProcessing();

            WriteDebug("Counting words and populating scaling dictionary.");
            return CountWords(_processedWords);
        }

        private void CompleteInputProcessing()
        {
            WriteDebug("Waiting for any remaining queued word processing work items to finish.");

            try
            {
                CancellableWaitAll(_waitHandles);
            }
            finally
            {
                WCUtils.DisposeAll(_waitHandles);
            }

            WriteDebug("Word processing tasks complete.");
        }

        private void CancellableWaitAll(IReadOnlyList<WaitHandle> handles)
        {
            var waitOrCancel = new WaitHandle[] { default!, _cancel.Token.WaitHandle };
            for (int index = 0; index < handles.Count; index++)
            {
                waitOrCancel[0] = handles[index];
                int waitHandleIndex = WaitHandle.WaitAny(waitOrCancel);
                if (waitHandleIndex == 1)
                {
                    // If we receive a signal from the cancellation token, throw PipelineStoppedException() to
                    // terminate the pipeline, as StopProcessing() has been called.
                    throw new PipelineStoppedException();
                }
            }
        }

        private static Dictionary<string, float> CountWords(IEnumerable<string> wordList)
        {
            var wordCounts = new Dictionary<string, float>(StringComparer.OrdinalIgnoreCase);

            foreach (string word in wordList)
            {
                var trimmedWord = Regex.Replace(word, "s$", string.Empty, RegexOptions.IgnoreCase);
                var pluralWord = string.Format("{0}s", word);
                if (wordCounts.ContainsKey(trimmedWord))
                {
                    wordCounts[trimmedWord]++;
                }
                else if (wordCounts.ContainsKey(pluralWord))
                {
                    wordCounts[word] = wordCounts[pluralWord] + 1;
                    wordCounts.Remove(pluralWord);
                }
                else
                {
                    wordCounts[word] = wordCounts.ContainsKey(word) ? wordCounts[word] + 1 : 1;
                }
            }

            return wordCounts;
        }
        private Dictionary<string, float> GetProcessedWordScaleDictionary(IDictionary dictionary)
        {
            var result = new Dictionary<string, float>(StringComparer.OrdinalIgnoreCase);

            WriteDebug("Processing -WordSizes input.");
            foreach (var word in dictionary.Keys)
            {
                if (word is null || WordSizes?[word] is null)
                {
                    continue;
                }

                if (TryConvertWordScale(word, WordSizes[word]!, out string wordString, out float wordScale))
                {
                    result[wordString] = wordScale;
                }
            }

            return result;
        }

        private bool TryConvertWordScale(object word, object scale, out string wordString, out float wordScale)
        {
            wordString = string.Empty;
            wordScale = 0;
            try
            {
                wordString = word.ConvertTo<string>();
                wordScale = scale.ConvertTo<float>();
                return true;
            }
            catch (Exception e)
            {
                WriteWarning($"Skipping entry '{word}' due to error converting key or value: {e.Message}.");
                WriteDebug($"Entry type: key - {word.GetType().FullName} ; value - {scale.GetType().FullName}");
                return false;
            }
        }

        IReadOnlyList<Word> GetWordListFromDictionary(IReadOnlyDictionary<string, float> dictionary, int maxWords)
        {
            IEnumerable<Word> sortedWords = dictionary
                .OrderByDescending(x => x.Value)
                .Select(x => new Word(text: x.Key, relativeSize: x.Value));

            if (MyInvocation.BoundParameters.ContainsKey(nameof(FocusWord)))
            {
                sortedWords = AddFocusWord(FocusWord!, sortedWords);
            }

            if (maxWords > 0)
            {
                sortedWords = sortedWords.Take(maxWords);
            }

            return sortedWords.ToList();
        }

        private IEnumerable<Word> AddFocusWord(string focusWord, IEnumerable<Word> words)
        {
            WriteDebug($"Adding focus word '{focusWord}' to the list.");

            float largestWordSize = words.Max(w => w.RelativeSize);
            return words.Prepend(new Word(focusWord, largestWordSize * Constants.FocusWordScale, isFocusWord: true));
        }

        #endregion

        #region Helpers - Drawing Words

        private Image CreateImage() => ParameterSetName.StartsWith(FILE_SET)
            ? new Image(BackgroundImage!, AllowOverflow.IsPresent)
            : new Image(ImageSize, BackgroundColor, AllowOverflow.IsPresent);

        private void DrawAllWordsOnCanvas(IReadOnlyList<Word> wordList, Image image)
        {
            WriteVerbose("Drawing words on canvas.");

            try
            {
                int totalWords = wordList.Count;
                float strokeWidth = MyInvocation.BoundParameters.ContainsKey(nameof(StrokeWidth))
                    ? StrokeWidth
                    : -1;

                for (int index = 0; index < wordList.Count; index++)
                {
                    Word currentWord = wordList[index];
                    WriteDebug($"Scanning for draw location for '{wordList[index]}'.");
                    WriteDrawingProgressMessage(
                        wordNumber: index,
                        totalWords,
                        currentWord.Text,
                        currentWord.ScaledSize);

                    DrawWord(currentWord, strokeWidth, image);
                }
            }
            finally
            {
                WCUtils.DisposeAll(wordList);
            }
        }

        private void DrawWord(Word word, float strokeWidth, Image image)
        {
            try
            {
                if (FindDrawLocation(word, Typeface, image))
                {
                    DrawWordInPlace(word, strokeWidth, image);
                }
                else
                {
                    WriteWarning($"Unable to find a place to draw '{word}'; skipping to next word.");
                }
            }
            catch (OperationCanceledException)
            {
                // If we receive OperationCancelledException, StopProcessing() has been called,
                // so the pipeline is terminating.
                throw new PipelineStoppedException();
            }
        }

        private bool FindDrawLocation(Word word, SKTypeface typeface, Image image)
        {
            WriteDebug($"Searching for an available point to draw '{word.Text}'");
            IReadOnlyList<float> availableAngles = word.IsFocusWord
                ? new[] { FocusWordAngle }
                : WCUtils.GetDrawAngles(AllowRotation, SafeRandom);

            using SKPaint brush = WCUtils.GetBrush(word.ScaledSize, StrokeWidth * Constants.StrokeBaseScale, typeface);
            word.Path = brush.GetTextPath(word.Text, 0, 0);
            word.Padding = GetPaddingValue(word);

            foreach (float drawAngle in availableAngles)
            {
                word.Path.Rotate(drawAngle);

                for (
                    float radius = SafeRandom.RandomFloat() / 25 * word.Path.TightBounds.Height;
                    radius <= image.MaxDrawRadius;
                    radius += GetRadiusIncrement(word, DistanceStep, image.MaxDrawRadius))
                {
                    if (FindDrawPointAtRadius(word, image, radius, drawAngle))
                    {
                        return true;
                    }
                }

                word.Path.Rotate(-drawAngle);
            }

            return false;
        }

        private bool FindDrawPointAtRadius(Word word, Image image, float radius, float drawAngle)
        {
            var ellipse = MyInvocation.BoundParameters.ContainsKey(nameof(RandomSeed))
                ? new Ellipse(radius, RadialStep, image, RandomSeed)
                : new Ellipse(radius, RadialStep, image);

            int totalPoints = ellipse.Points.Count;
            int pointsChecked = 0;

            foreach (SKPoint point in ellipse.Points)
            {
                ThrowIfPipelineStopping();

                pointsChecked++;
                if (point.IsOutside(image.ClippingRegion))
                {
                    WriteDebug($"Skipping point {point} because it's outside the clipping region.");
                    continue;
                }

                WritePointProgress(point, drawAngle, radius, pointsChecked, totalPoints);

                if (CanDrawWordUnobstructed(word, point, image))
                {
                    WriteDebug($"Found usable draw point at [{point.X}, {point.Y}]");
                    return true;
                }
            }

            return false;
        }

        private bool CanDrawWordUnobstructed(Word word, SKPoint point, Image image)
        {
            word.Path.CentreOnPoint(point);
            word.Bounds = SKRect.Inflate(word.Path.TightBounds, word.Padding, word.Padding);

            return WCUtils.WordWillFit(word, WordBubble, image);
        }

        private void DrawWordInPlace(Word word, float strokeWidth, Image image)
        {
            using var brush = new SKPaint();
            SKColor wordColor;
            SKColor bubbleColor;

            word.Path.FillType = SKPathFillType.Winding;
            if (word.Bubble is null)
            {
                wordColor = GetContrastingColor(BackgroundColor);
            }
            else
            {
                bubbleColor = GetContrastingColor(BackgroundColor);
                wordColor = GetContrastingColor(bubbleColor);

                brush.SetFill(bubbleColor);
                image.DrawPath(word.Bubble, brush);
            }

            brush.SetFill(wordColor);
            image.DrawPath(word.Path, brush);

            if (strokeWidth > -1)
            {
                brush.SetStroke(StrokeColor, strokeWidth);
                image.DrawPath(word.Path, brush);
            }
        }

        #endregion

        #region Helpers - Progress

        private void WriteDrawingProgressMessage(int wordNumber, int totalWords, string word, float wordSize)
        {
            var percentComplete = 100f * wordNumber / totalWords;

            var wordProgress = new ProgressRecord(
                _progressId,
                "Drawing word cloud...",
                $"Finding space for word: '{word}'...")
            {
                StatusDescription = string.Format(
                    "Current Word: \"{0}\" [Size: {1:0}] ({2} of {3})",
                    word,
                    wordSize,
                    wordNumber,
                    totalWords),
                PercentComplete = (int)Math.Round(percentComplete)
            };

            WriteProgress(wordProgress);
        }

        private void WriteProgressCompleted()
        {
            WriteProgress(new ProgressRecord(_progressId, "Completed", "Completed")
            {
                RecordType = ProgressRecordType.Completed
            });

            WriteProgress(new ProgressRecord(_progressId + 1, "Completed", "Completed")
            {
                RecordType = ProgressRecordType.Completed
            });
        }

        private void WritePointProgress(SKPoint point, float drawAngle, float radius, int pointsChecked, int totalPoints)
        {
            var pointProgress = new ProgressRecord(
                _progressId + 1,
                "Scanning available space...",
                "Scanning radial points...")
            {
                ParentActivityId = _progressId,
                Activity = string.Format(
                    "Finding available space to draw at angle: {0}",
                    drawAngle),
                StatusDescription = string.Format(
                    "Checking [Point:{0,8:N2}, {1,8:N2}] ({2,4} / {3,4}) at [Radius: {4,8:N2}]",
                    point.X,
                    point.Y,
                    pointsChecked,
                    totalPoints,
                    radius)
            };

            WriteProgress(pointProgress);
        }

        #endregion

        #region Helpers - Errors

        private void ThrowIfPipelineStopping() => _cancel.Token.ThrowIfCancellationRequested();

        private void ThrowIfEmpty<T>(IReadOnlyList<T> items)
        {
            if (items.Count == 0)
            {
                ThrowTerminatingError(
                    new ArgumentException("No usable input was provided. Please provide string data via the pipeline or -WordSizes."),
                    ErrorCodes.NoUsableInput,
                    ErrorCategory.InvalidData,
                    MyInvocation.BoundParameters.ContainsKey(nameof(InputObject))
                        ? InputObject?.BaseObject
                        : WordSizes);
            }
        }

        private void ThrowTerminatingError(
            Exception exception,
            string code,
            ErrorCategory category,
            object? target)
        {
            var errorRecord = new ErrorRecord(exception, code, category, target);
            ThrowTerminatingError(errorRecord);
        }

        #endregion

        #region Helpers - File Handling

        private void SaveSvgData(Image image, string savePath)
        {
            string[] path = new[] { savePath };

            ClearItemContent(path);
            WriteSvgDataToPSPath(path, image.GetFinalXml());
        }

        private void ClearItemContent(string[] paths)
        {
            WriteDebug($"Clearing existing content from '{string.Join(", ", paths)}'.");
            for (int index = 0; index < paths.Length; index++)
            {
                ClearOrRemoveExistingItem(paths[index]);
            }
        }

        private void ClearOrRemoveExistingItem(string path)
        {
            if (!InvokeProvider.Item.Exists(path, force: false, literalPath: true))
            {
                return;
            }

            string[] target = new[] { path };
            try
            {
                WriteDebug($"Item '{path}' already exists, clearing existing contents.");
                InvokeProvider.Content.Clear(target, force: false, literalPath: true);
            }
            catch (Exception e)
            {
                WriteDebug($"An error was encountered attempting to clear the content for '{path}': {e.Message}");
                WriteDebug($"Attempting to remove the existing item at '{path}' instead...");

                InvokeProvider.Item.Remove(target, recurse: false, force: false, literalPath: true);
            }
        }

        private void WriteSvgDataToPSPath(string[] paths, XmlDocument svgData)
        {
            WriteDebug($"Saving data to '{Path}'.");
            using IContentWriter writer = InvokeProvider.Content.GetWriter(paths, force: false, literalPath: true).First();
            writer.Write(new[] { svgData.GetPrettyString() });
            writer.Close();
        }

        #endregion

        #region Helpers - Miscellaneous

        private float GetPaddingValue(Word word)
        {
            float padding = word.Path.TightBounds.Height * PaddingMultiplier / word.ScaledSize;
            return padding;
        }

        private IReadOnlyList<Word> ScaleWordSizes(
            IReadOnlyList<Word> wordList,
            float maxWordWidth,
            float baseFontScale,
            float userFontScale)
        {
            using SKPaint brush = WCUtils.GetBrush(wordSize: 0, StrokeWidth * Constants.StrokeBaseScale, Typeface);
            for (int index = 0; index < wordList.Count; index++)
            {
                brush.TextSize = wordList[index].Scale(baseFontScale * userFontScale);
                wordList[index].Path = brush.GetTextPath(wordList[index].Text, x: 0, y: 0);
                SKRect textRect = wordList[index].Bounds;

                float paddedWidth = textRect.Width + GetPaddingValue(wordList[index]);
                if (!AllowOverflow.IsPresent && paddedWidth > maxWordWidth)
                {
                    return ScaleWordSizes(
                        wordList,
                        maxWordWidth,
                        baseFontScale * Constants.MaxWordWidthPercent * (maxWordWidth / paddedWidth),
                        userFontScale);
                }
            }

            return wordList;
        }

        private SKColor GetNextColor()
        {
            if (_colorSetIndex >= ColorSet.Length)
            {
                _colorSetIndex = 0;
            }

            return ColorSet[_colorSetIndex++];
        }

        private SKColor GetContrastingColor(SKColor reference)
        {
            SKColor result;
            uint attempts = 0;
            do
            {
                result = GetNextColor();
                attempts++;
            }
            while (!result.IsDistinctFrom(reference) && attempts < ColorSet.Length);

            return result;
        }

        private static float GetRadiusIncrement(Word word, float distanceStep, float maxRadius)
        {
            var wordScaleFactor = (1 + word.Padding) * word.ScaledSize / 360;
            var stepScale = distanceStep / 15;
            var minRadiusIncrement = maxRadius / 1000;

            var radiusIncrement = minRadiusIncrement * stepScale + wordScaleFactor;

            return radiusIncrement;
        }

        #endregion
    }
}
