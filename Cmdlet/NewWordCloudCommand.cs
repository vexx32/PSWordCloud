using System;
using System.Collections.Generic;
using System.Management.Automation;
using System.Threading.Tasks;
using SkiaSharp;

namespace PSWordCloud
{
    [Cmdlet(VerbsCommon.New, "WordCloud", SupportsShouldProcess = true, DefaultParameterSetName = "ColorBackground")]
    public class NewWordCloudCommand : PSCmdlet
    {
        #region Parameters

        [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "ColorBackground")]
        [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "ColorBackground-Mono")]
        [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "FileBackground")]
        [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "FileBackground-Mono")]
        [Alias("InputString", "Text", "String", "Words", "Document", "Page")]
        [AllowEmptyString()]
        public PSObject InputObject { get; set; }

        private string[] _paths;
        [Parameter(Mandatory = true, Position = 1, ParameterSetName = "ColorBackground")]
        [Parameter(Mandatory = true, Position = 1, ParameterSetName = "ColorBackground-Mono")]
        [Parameter(Mandatory = true, Position = 1, ParameterSetName = "FileBackground")]
        [Parameter(Mandatory = true, Position = 1, ParameterSetName = "FileBackground-Mono")]
        [Alias("OutFile", "ExportPath", "ImagePath")]
        public string[] Path { get; set; }

        #endregion Parameters

        private List<string> _inputCache = new List<string>(256);
        private List<Task<string[]>> _taskCache = new List<Task<string[]>>();
        private static string[] _stopWords = {
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
            "yours","yourself","yourselves"
        }

        protected override void BeginProcessing()
        {
            var targetPaths = new List<string>();

            foreach (string path in Path)
            {
                var resolvedPaths = SessionState.Path.GetResolvedProviderPathFromPSPath(path, out ProviderInfo provider);
                if (resolvedPaths != null)
                {
                    targetPaths.AddRange(resolvedPaths);
                }
            }

            _paths = targetPaths.ToArray();
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

            _inputCache.AddRange(text);

            if (!MyInvocation.ExpectingInput || _inputCache.Count >= 250)
            {
                // Kick off async job
            }
        }

        protected override void EndProcessing()
        {

        }

        private async Task<string[]> SubdivideTextAsync(string[] lines)
        {
            List<string> words = new List<string>(lines.Length);



            return words.ToArray();
        }
    }
}
