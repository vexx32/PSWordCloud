using System;
using System.Collections.Generic;
using System.Management.Automation;
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
        public string Path { get; set; }

        [Parameter(ParameterSetName = "ColorBackground")]
        [Parameter(ParameterSetName = "ColorBackground-Mono")]
        [Parameter(ParameterSetName = "FileBackground")]
        [Parameter(ParameterSetName = "FileBackground-Mono")]
        [Alias("ColourSet")]
        [ArgumentCompleter(ColorCompleter)]
        [ColorTransformAttribute()]
        [Color[]]
        public object[] ColorSet = [ColorTransformAttribute]::ColorNames

        #endregion Parameters

        protected override void BeginProcessing()
        {
            List<string> targetPaths = new List<string>();

            foreach (string path in Path)
            {
                var resolvedPaths = SessionState.Path.GetResolvedProviderPathFromPSPath(path, out ProviderInfo provider);
                if (resolvedPaths != null)
                {
                    targetPaths.AddRange(resolvedPaths);
                }
            }

            _paths = targetPaths.ToString();
        }

        protected override void ProcessRecord()
        {

            WriteObject(_paths, true);
        }

        protected override void EndProcessing()
        {

        }


    }
}
