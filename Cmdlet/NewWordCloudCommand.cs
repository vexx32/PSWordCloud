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
        [ToSKSizeITransform]
        public SKSizeI ImageSize { get; set; } = new SKSizeI(4096, 2304);

        [Parameter]
        [ArgumentCompleter(typeof(FontFamilyCompleter))]
        [ToSKTypefaceTransform]
        public SKTypeface FontFamily { get; set; } = SKTypeface.FromFamilyName(
            "Consolas",
            SKFontStyleWeight.Normal,
            SKFontStyleWidth.Normal,
            SKFontStyleSlant.Upright);

        [Parameter]
        [Alias("Title")]
        public string FocusWord { get; set; }

        [Parameter]
        [Alias("ScaleFactor")]
        [ValidateRange(0.01, 5)]
        public float WordScale { get; set; } = 1f;

        [Parameter]
        [Alias("MaxWords")]
        [ValidateRange(0, 1000)]
        public ushort MaxRenderedWords { get; set; } = 100;

        [Parameter]
        [Alias("SeedValue")]
        public int RandomSeed { get; set; }

        [Parameter]
        [Alias("AllowOverflow")]
        public SwitchParameter AllowBleed { get; set; }

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

        protected override void BeginProcessing()
        {
            _random = MyInvocation.BoundParameters.ContainsKey("RandomSeed") ? new Random(RandomSeed) : new Random();

            var targetPaths = new List<string>();

            foreach (string path in Path)
            {
                var resolvedPaths = SessionState.Path.GetResolvedProviderPathFromPSPath(path, out ProviderInfo provider);
                if (resolvedPaths != null)
                {
                    targetPaths.AddRange(resolvedPaths);
                }
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
            Dictionary<string, float> wordScaleDictionary = new Dictionary<string, float>(
                StringComparer.OrdinalIgnoreCase);

            var lineStrings = Task.WhenAll<string[]>(_wordProcessingTasks);
            lineStrings.Wait();

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
            var lowestWordFreq = wordSizeValues.Min();
            if (FocusWord != null)
            {
                wordScaleDictionary[FocusWord] = highestWordFreq = highestWordFreq * FOCUS_WORD_SCALE;
            }

            float averageWordFrequency = wordSizeValues.Average();

            string[] sortedWordList = wordScaleDictionary.Keys
                .OrderByDescending(size => wordScaleDictionary[size])
                .Take(MaxRenderedWords == 0 ? ushort.MaxValue : MaxRenderedWords)
                .ToArray();

            try
            {
                SKRectI bounds = new SKRectI(0, 0, ImageSize.Width, ImageSize.Height);
                if (AllowBleed.IsPresent)
                {
                    bounds.Inflate((int)(bounds.Width * BLEED_AREA_SCALE), (int)(bounds.Height * BLEED_AREA_SCALE));
                }

                float fontScale = WordScale * 1.6f *
                        (bounds.Height + bounds.Width) / (averageWordFrequency * sortedWordList.Length);


                SKRegion occupiedRegion = new SKRegion();
                SKRectI barrierExtent = SKRectI.Inflate(bounds, bounds.Width * 2, bounds.Height * 2);
                occupiedRegion.SetRect(barrierExtent);
                occupiedRegion.Op(bounds, SKRegionOperation.Difference);

                Dictionary<string, float> finalWordSizes = new Dictionary<string, float>(
                    sortedWordList.Length, StringComparer.OrdinalIgnoreCase);

                foreach (string word in sortedWordList)
                {
                    var finalWordSize = (float)Math.Round(
                        2 * wordScaleDictionary[word] * fontScale * _random.NextDouble() /
                        (1f + highestWordFreq - lowestWordFreq) + 0.9);

                    if (finalWordSize < 5) continue;


                }

                // Basic SVG canvas creation
                SKFileWStream streamWriter = new SKFileWStream(_resolvedPaths[0]);
                SKXmlStreamWriter xmlWriter = new SKXmlStreamWriter(streamWriter);
                SKCanvas canvas = SKSvgCanvas.Create(bounds, xmlWriter);

                // TODO: Ensure saved file is copied to all _resolvedPaths
            }
            catch (Exception e)
            {
                WriteError(new ErrorRecord(e, "PSWordCloudError", ErrorCategory.InvalidResult, null));
            }
            finally
            {

            }
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
