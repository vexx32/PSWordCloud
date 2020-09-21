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

        #endregion Constants

        #region StaticMembers

        private static LockingRandom? _lockingRandom;
        internal static LockingRandom SafeRandom => _lockingRandom ??= new LockingRandom();

        private readonly CancellationTokenSource _cancel = new CancellationTokenSource();

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

        #endregion Parameters

        #region privateVariables

        private float PaddingMultiplier { get => Padding * Constants.PaddingBaseScale; }

        private readonly ConcurrentBag<string> _processedWords = new ConcurrentBag<string>();
        private readonly List<EventWaitHandle> _waitHandles = new List<EventWaitHandle>();

        private int _colorSetIndex = 0;

        private int _progressId;

        #endregion privateVariables

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

        /// <summary>
        /// Implements the <see cref="ProcessRecord"/> method for <see cref="NewWordCloudCommand"/>.
        /// Spins up a <see cref="Task{IEnumerable{string}}"/> for each input text string to split them all
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

        private void QueueInputProcessing(PSObject inputObject)
        {
            var state = new EventWaitHandle(initialState: false, EventResetMode.ManualReset);
            ThreadPool.QueueUserWorkItem(StartProcessingInput, (inputObject, state), preferLocal: false);
            _waitHandles.Add(state);
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

        private IEnumerable<Word> AddFocusWord(string focusWord, IEnumerable<Word> words)
        {
            WriteDebug($"Adding focus word '{focusWord}' to the list.");

            float largestWordSize = words.Max(w => w.RelativeSize);
            return words.Prepend(new Word(focusWord, largestWordSize * Constants.FocusWordScale, isFocusWord: true));
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

        private SKRect GetImageViewbox(string? backgroundImagePath, out SKBitmap? backgroundBitmap)
        {
            backgroundBitmap = null;
            if (backgroundImagePath is null)
            {
                return new SKRect(left: 0, top: 0, ImageSize.Width, ImageSize.Height);
            }
            else
            {
                backgroundBitmap = LoadBackground(backgroundImagePath, out SKRect backgroundRect);
                return backgroundRect;
            }
        }

        private SKBitmap LoadBackground(string path, out SKRect backgroundRect)
        {
            WriteDebug($"Importing background image from '{path}'.");
            SKBitmap backgroundBitmap = SKBitmap.Decode(path);
            backgroundRect = new SKRectI(left: 0, top: 0, backgroundBitmap.Width, backgroundBitmap.Height);
            return backgroundBitmap;
        }

        private float GetMaxWordWidth(SKRect viewbox, float focusWordAngle)
            => WCUtils.AngleIsMostlyVertical(focusWordAngle)
                ? viewbox.Height * Constants.MaxWordWidthPercent
                : viewbox.Width * Constants.MaxWordWidthPercent;

        private float GetMaxWordWidth(SKRect viewbox, WordOrientations permittedOrientations)
            => permittedOrientations == WordOrientations.None
                ? viewbox.Width * Constants.MaxWordWidthPercent
                : Math.Max(viewbox.Width, viewbox.Height) * Constants.MaxWordWidthPercent;

        private void DrawImageBackground(
            Image image,
            SKColor backgroundColor = default,
            SKBitmap? backgroundImage = null)
        {
            if (ParameterSetName.StartsWith(FILE_SET))
            {
                image.Canvas.DrawBitmap(backgroundImage, 0, 0);
            }
            else if (backgroundColor != SKColor.Empty)
            {
                image.Canvas.Clear(backgroundColor);
            }
        }

        private void DrawAllWordsOnCanvas(IReadOnlyList<Word> wordList, Image image)
        {
            WriteVerbose("Drawing words on canvas.");

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

        private void DrawWord(Word word, float strokeWidth, Image image)
        {
            SKPath? wordPath = null, bubblePath = null;
            try
            {
                if (FindDrawLocation(word, Typeface, image, out wordPath, out bubblePath))
                {
                    DrawWordInPlace(wordPath, strokeWidth, bubblePath, image);
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
            finally
            {
                wordPath?.Dispose();
                bubblePath?.Dispose();
            }
        }

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

        private (float averageFrequency, float averageLength) GetWordListStatistics(
            IReadOnlyList<Word> wordList)
        {

            float totalFrequency = 0, totalLength = 0;
            for (int i = 0; i < wordList.Count; i++)
            {
                totalFrequency += wordList[i].RelativeSize;
                totalLength += wordList[i].Text.Length;
            }

            return (totalFrequency / wordList.Count, totalLength / wordList.Count);
        }

        private float GetPaddingValue(Word word, SKPath wordPath)
        {
            float padding = wordPath.TightBounds.Height * PaddingMultiplier / word.ScaledSize;
            return padding;
        }

        private float ConstrainScaleByWordWidth(
            float maxWordWidth,
            Word largestWord,
            float fontScale)
        {
            float largestWordSize = largestWord.Scale(fontScale);

            using SKPaint brush = WCUtils.GetBrush(largestWordSize, StrokeWidth * Constants.StrokeBaseScale, Typeface);
            using SKPath wordPath = brush.GetTextPath(largestWord.Text, x: 0, y: 0);

            float padding = GetPaddingValue(largestWord, wordPath);
            float effectiveWidth = Math.Max(Constants.MinEffectiveWordWidth, largestWord.Text.Length);
            float adjustedWidth = wordPath.Bounds.Width * (effectiveWidth / largestWord.Text.Length) + padding;
            if (adjustedWidth > maxWordWidth)
            {
                return ConstrainScaleByWordWidth(
                    maxWordWidth,
                    largestWord,
                    fontScale * Constants.MaxWordWidthPercent * (maxWordWidth / adjustedWidth));
            }

            return fontScale;
        }

        private static float EstimateWordScale(
            SKRect canvasRect,
            float averageWordFrequency,
            float averageWordLength,
            int wordCount,
            SKTypeface typeface)
        {
            float fontCharArea = WCUtils.GetAverageCharArea(typeface);
            float estimatedPadding = (float)Math.Sqrt(fontCharArea) * Constants.PaddingBaseScale / averageWordFrequency;
            float estimatedWordArea = fontCharArea * averageWordLength * averageWordFrequency * wordCount + estimatedPadding;
            float canvasArea = canvasRect.Height * canvasRect.Width;

            return canvasArea * Constants.MaxWordAreaPercent / estimatedWordArea;
        }

        private float GetWordScaleFactor(IReadOnlyList<Word> wordList, float maxWordWidth, SKRect drawableBounds)
        {
            (float averageWordFrequency, float averageWordLength) = GetWordListStatistics(wordList);

            float estimatedScale = EstimateWordScale(
                drawableBounds,
                averageWordFrequency,
                averageWordLength,
                wordList.Count,
                Typeface);

            return ConstrainScaleByWordWidth(maxWordWidth, wordList[0], estimatedScale);
        }

        private IReadOnlyList<Word> GetFinalWordList(SKRectI drawableBounds, SKRect viewbox)
        {
            IReadOnlyList<Word> wordList = GetRelativeWordSizes(ParameterSetName);
            if (wordList.Count == 0)
            {
                ThrowTerminatingError(
                    new ErrorRecord(
                        new ArgumentException("No usable input was provided. Please provide string data via the pipeline or in a word size dictionary."),
                        ErrorCodes.NoUsableInput,
                        ErrorCategory.InvalidData,
                        MyInvocation.BoundParameters.ContainsKey(nameof(InputObject))
                            ? InputObject?.BaseObject
                            : WordSizes));
            }

            float maxWordWidth = MyInvocation.BoundParameters.ContainsKey(nameof(FocusWord))
                ? GetMaxWordWidth(viewbox, FocusWordAngle)
                : GetMaxWordWidth(viewbox, AllowRotation);

            float scaleFactor = GetWordScaleFactor(wordList, maxWordWidth, drawableBounds);

            WriteDebug($"Global font scale: {scaleFactor}");

            return ScaleWordSizes(
                wordList,
                maxWordWidth,
                scaleFactor,
                WordScale);
        }

        private void CreateWordCloud()
        {
            SKRect viewbox = GetImageViewbox(BackgroundImage, out SKBitmap? backgroundBitmap);
            if (backgroundBitmap is not null)
            {
                BackgroundColor = WCUtils.GetAverageColor(backgroundBitmap.Pixels);
            }

            using var image = new Image(viewbox, AllowOverflow.IsPresent);

            IReadOnlyList<Word> finalWordTable = GetFinalWordList(image.ClippingBounds, viewbox);

            try
            {
                DrawImageBackground(image, BackgroundColor, backgroundBitmap);
                DrawAllWordsOnCanvas(finalWordTable, image);

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
                WriteDebug("Disposing SkiaSharp objects.");
                backgroundBitmap?.Dispose();

                WriteProgressCompleted();
            }
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

        /// <summary>
        /// StopProcessing implementation for New-WordCloud.
        /// </summary>
        protected override void StopProcessing()
        {
            // Cancellation is registered in both the word-processing and word placement code regions,
            // which comprise the majority of the "slow" / "blocking" behaviour of the cmdlet.
            _cancel.Cancel();
        }

        #region HelperMethods

        private void DrawWordInPlace(SKPath wordPath, float strokeWidth, SKPath? bubblePath, Image image)
        {
            using var brush = new SKPaint();
            SKColor wordColor;
            SKColor bubbleColor;

            wordPath.FillType = SKPathFillType.Winding;
            if (bubblePath == null)
            {
                wordColor = GetContrastingColor(BackgroundColor);
            }
            else
            {
                bubbleColor = GetContrastingColor(BackgroundColor);
                wordColor = GetContrastingColor(bubbleColor);

                brush.SetFill(bubbleColor);
                image.DrawPath(bubblePath, brush);
            }

            brush.SetFill(wordColor);
            image.DrawPath(wordPath, brush);

            if (strokeWidth > -1)
            {
                brush.SetStroke(StrokeColor, strokeWidth);
                image.DrawPath(wordPath, brush);
            }
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
                using SKPath textPath = brush.GetTextPath(wordList[index].Text, x: 0, y: 0);
                SKRect textRect = textPath.Bounds;

                float paddedWidth = textRect.Width + GetPaddingValue(wordList[index], textPath);
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

        private bool FindDrawLocation(
            Word word,
            SKTypeface typeface,
            Image image,
            out SKPath wordPath,
            out SKPath? bubblePath)
        {
            WriteDebug($"Searching for an available point to draw '{word.Text}'");
            IReadOnlyList<float> availableAngles = word.IsFocusWord
                ? new[] { FocusWordAngle }
                : WCUtils.GetDrawAngles(AllowRotation, SafeRandom);

            bubblePath = null;

            using SKPaint brush = WCUtils.GetBrush(word.ScaledSize, StrokeWidth * Constants.StrokeBaseScale, typeface);
            wordPath = brush.GetTextPath(word.Text, 0, 0);

            float wordPadding = GetPaddingValue(word, wordPath);

            foreach (float drawAngle in availableAngles)
            {
                wordPath.Rotate(drawAngle);

                for (
                    float radius = SafeRandom.RandomFloat() / 25 * wordPath.TightBounds.Height;
                    radius <= image.MaxDrawRadius;
                    radius += GetRadiusIncrement(
                        word.ScaledSize,
                        DistanceStep,
                        image.MaxDrawRadius,
                        wordPadding))
                {
                    if (FindDrawPointAtRadius(
                        image,
                        radius,
                        drawAngle,
                        wordPadding,
                        wordPath,
                        out bubblePath))
                    {
                        return true;
                    }
                }

                wordPath.Rotate(-drawAngle);
            }

            return false;
        }

        private void ThrowIfPipelineStopping() => _cancel.Token.ThrowIfCancellationRequested();

        private bool FindDrawPointAtRadius(
            Image image,
            float radius,
            float drawAngle,
            float inflationValue,
            SKPath wordPath,
            out SKPath? bubblePath)
        {
            IReadOnlyList<SKPoint> radialPoints = GetOvalPoints(radius, RadialStep, image);
            int totalPoints = radialPoints.Count;
            int pointsChecked = 0;
            bubblePath = null;

            foreach (SKPoint point in radialPoints)
            {
                ThrowIfPipelineStopping();

                pointsChecked++;
                if (point.IsOutside(image.ClippingRegion))
                {
                    WriteDebug($"Skipping point {point} because it's outside the clipping region.");
                    continue;
                }

                WritePointProgress(point, drawAngle, radius, pointsChecked, totalPoints);

                wordPath.CentreOnPoint(point);
                SKRect wordBounds = SKRect.Inflate(wordPath.TightBounds, inflationValue, inflationValue);

                if (WCUtils.WordWillFit(wordBounds, WordBubble, image, out bubblePath))
                {
                    WriteDebug($"Found usable draw point at [{point.X}, {point.Y}]");
                    return true;
                }
            }

            return false;
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

        private void SaveSvgData(Image image, string savePath)
        {
            string[] path = new[] { savePath };

            ClearFileContent(path);
            WriteSvgDataToPSPath(path, image.GetFinalXml());
        }

        private void ClearFileContent(string[] paths)
        {
            try
            {
                WriteDebug($"Clearing existing content from '{string.Join(", ", paths)}'.");
                InvokeProvider.Content.Clear(paths, force: false, literalPath: true);
            }
            catch (Exception e)
            {
                // Unconditionally suppress errors from the Content.Clear() operation. Errors here may indicate that
                // a provider is being written to that does not support the Content.Clear() interface, so ignore errors
                // at this point.
                WriteDebug($"Error encountered while clearing content for item '{string.Join(", ", paths)}'. {e.Message}");
            }
        }

        private void WriteSvgDataToPSPath(string[] paths, XmlDocument svgData)
        {
            WriteDebug($"Saving data to '{Path}'.");
            using IContentWriter writer = InvokeProvider.Content.GetWriter(paths, force: false, literalPath: true).First();
            writer.Write(new[] { svgData.GetPrettyString() });
            writer.Close();
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

            if (enumerable != null)
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

        private static float GetRadiusIncrement(
            float wordSize,
            float distanceStep,
            float maxRadius,
            float padding)
        {
            var wordScaleFactor = (1 + padding) * wordSize / 360;
            var stepScale = distanceStep / 15;
            var minRadiusIncrement = maxRadius / 1000;

            var radiusIncrement = minRadiusIncrement * stepScale + wordScaleFactor;

            return radiusIncrement;
        }

        private static IReadOnlyList<SKPoint> GetOvalPoints(float radius, float radialStep, Image image)
        {
            var result = new List<SKPoint>();
            if (radius == 0)
            {
                result.Add(image.Centre);
                return result;
            }

            float angleIncrement = radialStep * Constants.BaseAngularIncrement / (15 * (float)Math.Sqrt(radius));

            float angle = SafeRandom.PickRandomQuadrant();
            bool clockwise = SafeRandom.RandomFloat() > 0.5;

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

            return GenerateOvalPoints(
                radius,
                angle,
                angleIncrement,
                maxAngle,
                image);
        }

        private static IReadOnlyList<SKPoint> GenerateOvalPoints(
            float radius,
            float startingAngle,
            float angleIncrement,
            float maxAngle,
            Image image)
        {
            List<SKPoint> points = new List<SKPoint>();
            float angle = startingAngle;
            bool clockwise = angleIncrement > 0;

            do
            {
                points.Add(GetPointOnOval(radius, angle, image));
                angle += angleIncrement;
            } while (clockwise ? angle <= maxAngle : angle >= maxAngle);

            return points;
        }

        private static SKPoint GetPointOnOval(float semiMinorAxis, float degrees, Image image)
        {
            Complex point = Complex.FromPolarCoordinates(semiMinorAxis, degrees.ToRadians());
            float xPosition = image.Centre.X + (float)point.Real * image.AspectRatio;
            float yPosition = image.Centre.Y + (float)point.Imaginary;

            return new SKPoint(xPosition, yPosition);
        }

        private void CompleteInputProcessing()
        {
            WriteDebug("Waiting for any remaining queued word processing work items to finish.");

            var waitHandles = new WaitHandle[] { default!, _cancel.Token.WaitHandle };
            try
            {
                for (int index = 0; index < _waitHandles.Count; index++)
                {
                    waitHandles[0] = _waitHandles[index];
                    int waitHandleIndex = WaitHandle.WaitAny(waitHandles);
                    if (waitHandleIndex == 1)
                    {
                        // If we receive a signal from the cancellation token, throw PipelineStoppedException() to
                        // terminate the pipeline, as StopProcessing() has been called.
                        throw new PipelineStoppedException();
                    }
                }
            }
            finally
            {
                WCUtils.DisposeAll(_waitHandles);
            }

            WriteDebug("Word processing tasks complete.");
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

        private Dictionary<string, float> GetWordCountDictionary()
        {
            CompleteInputProcessing();

            WriteDebug("Counting words and populating scaling dictionary.");
            return CountWords(_processedWords);
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

            if (Regex.Replace(word, "[^a-zA-Z-]", string.Empty).Length < 2)
            {
                return false;
            }

            return true;
        }

        #endregion HelperMethods
    }
}
