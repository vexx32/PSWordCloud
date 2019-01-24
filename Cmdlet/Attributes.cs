using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Language;
using System.Text.RegularExpressions;
using SkiaSharp;


namespace PSWordCloud
{
    public class ImageSizeCompleter : IArgumentCompleter
    {
        public IEnumerable<CompletionResult> CompleteArgument(
            string commandName,
            string parameterName,
            string wordToComplete,
            CommandAst commandAst,
            IDictionary fakeboundParameters)
        {

            var matchingResults = SKSizeTransform.StandardImageSizes.Where(
                keyPair => keyPair.Key.StartsWith(wordToComplete, StringComparison.OrdinalIgnoreCase));

            foreach (KeyValuePair<string, (string Tooltip, SKSize)> result in matchingResults)
            {
                yield return new CompletionResult(
                    result.Key,
                    result.Key,
                    CompletionResultType.ParameterValue,
                    result.Value.Tooltip);
            }
        }
    }

    public class SKSizeTransform : ArgumentTransformationAttribute
    {
        public static Dictionary<string, (string Tooltip, SKSize Size)> StandardImageSizes =
        new Dictionary<string, (string, SKSize)>() {
                {"480x800",     ("Mobile Screen Size (small)",  new SKSize(480, 800)  )},
                {"640x1146",    ("Mobile Screen Size (medium)", new SKSize(640, 1146) )},
                {"720p",        ("Standard HD 1280x720",        new SKSize(1280, 720) )},
                {"1080p",       ("Full HD 1920x1080",           new SKSize(1920, 1080))},
                {"4K",          ("Ultra HD 3840x2160",          new SKSize(3840, 2160))},
                {"A4",          ("816x1056",                    new SKSize(816, 1056) )},
                {"Poster11x17", ("1056x1632",                   new SKSize(1056, 1632))},
                {"Poster18x24", ("1728x2304",                   new SKSize(1728, 2304))},
                {"Poster24x36", ("2304x3456",                   new SKSize(2304, 3456))},
            };
        public override object Transform(EngineIntrinsics engineIntrinsics, object inputData)
        {
            float sideLength = 0;
            switch (inputData)
            {
                case short sh:
                    sideLength = sh;
                    break;
                case ushort us:
                    sideLength = us;
                    break;
                case int i:
                    sideLength = i;
                    break;
                case uint u:
                    sideLength = u;
                    break;
                case long l:
                    sideLength = l;
                    break;
                case ulong ul:
                    sideLength = ul;
                    break;
                case decimal d:
                    sideLength = (float)d;
                    break;
                case float f:
                    sideLength = f;
                    break;
                case double d:
                    if (d <= float.MaxValue)
                    {
                        sideLength = (float)d;
                        break;
                    }
                    else
                    {
                        throw new ArgumentTransformationMetadataException("Value too large for image size");
                    }
                case string s:
                    if (StandardImageSizes.ContainsKey(s))
                    {
                        return StandardImageSizes[s].Size;
                    }
                    else
                    {
                        var matchWH = Regex.Match(s, @"^(?<Width>[\d\.,]+)x(?<Height>[\d\.,]+)(px)?$");
                        if (matchWH.Success)
                        {
                            try
                            {
                                var width = float.Parse(matchWH.Groups["Width"].Value);
                                var height = float.Parse(matchWH.Groups["Height"].Value);

                                return new SKSize(width, height);
                            }
                            catch (Exception e)
                            {
                                throw new ArgumentTransformationMetadataException(
                                    "Could not parse input string as a float value", e);
                            }
                        }

                        var matchSide = Regex.Match(s, @"^(?<SideLength>[\d\.,]+)(px)?$");
                        if (matchSide.Success)
                        {
                            sideLength = float.Parse(matchSide.Groups["SideLength"].Value);
                        }
                    }

                    break;
                default:
                    throw new ArgumentTransformationMetadataException();
            }

            if (sideLength > 0)
            {
                return new SKSize(sideLength, sideLength);
            }

            throw new ArgumentTransformationMetadataException();
        }
    }
}