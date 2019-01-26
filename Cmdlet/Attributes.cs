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

            var matchingResults = WordCloudUtils.StandardImageSizes.Where(
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

    public class ToSKSizeITransformAttribute : ArgumentTransformationAttribute
    {
        public override object Transform(EngineIntrinsics engineIntrinsics, object inputData)
        {
            dynamic sideLength = 0;
            switch (inputData)
            {
                case SKSize sk:
                    return sk;
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
                    sideLength = d;
                    break;
                case float f:
                    sideLength = f;
                    break;
                case double d:
                    sideLength = d;
                    break;
                case string s:
                    if (WordCloudUtils.StandardImageSizes.ContainsKey(s))
                    {
                        return WordCloudUtils.StandardImageSizes[s].Size;
                    }
                    else
                    {
                        var matchWH = Regex.Match(s, @"^(?<Width>[\d\.,]+)x(?<Height>[\d\.,]+)(px)?$");
                        if (matchWH.Success)
                        {
                            try
                            {
                                var width = int.Parse(matchWH.Groups["Width"].Value);
                                var height = int.Parse(matchWH.Groups["Height"].Value);

                                return new SKSizeI(width, height);
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
                            sideLength = int.Parse(matchSide.Groups["SideLength"].Value);
                        }
                    }

                    break;
                case object o:
                    PSObject ps = PSObject.AsPSObject(o);
                    if (ps.Properties["Width"] != null && ps.Properties["Height"] != null)
                    {
                        // If these conversions fail, the exception will cause the transform to fail.
                        var width = LanguagePrimitives.ConvertTo<int>(ps.Properties["Width"]);
                        var height = LanguagePrimitives.ConvertTo<int>(ps.Properties["Height"]);

                        return new SKSizeI(width, height);
                    }

                    break;
                default:
                    throw new ArgumentTransformationMetadataException();
            }

            if (sideLength > 0 && sideLength <= int.MaxValue)
            {
                return new SKSizeI((int)sideLength, (int)sideLength);
            }

            throw new ArgumentTransformationMetadataException();
        }
    }

    public class FontFamilyCompleter : IArgumentCompleter
    {
        private static List<string> _fontList = new List<string>(WordCloudUtils.FontManager.GetFontFamilies());

        public IEnumerable<CompletionResult> CompleteArgument(
            string commandName,
            string parameterName,
            string wordToComplete,
            CommandAst commandAst,
            IDictionary fakeBoundParameters)
        {
            var fonts = WordCloudUtils.FontManager.GetFontFamilies();
            if (string.IsNullOrEmpty(wordToComplete))
            {
                foreach (string font in _fontList)
                {
                    yield return new CompletionResult(font, font, CompletionResultType.ParameterValue, null);
                }
            }
            else
            {
                foreach (string font in _fontList.Where(
                    s => s.StartsWith(wordToComplete, StringComparison.CurrentCultureIgnoreCase)))
                {
                    yield return new CompletionResult(font, font, CompletionResultType.ParameterName, null);
                }
            }
        }
    }

    public class ToSKTypefaceTransformAttribute : ArgumentTransformationAttribute
    {
        public override object Transform(EngineIntrinsics engineIntrinsics, object inputData)
        {
            switch (inputData)
            {
                case SKTypeface t:
                    return t;
                case string s:
                    return WordCloudUtils.FontManager.MatchFamily(s, SKFontStyle.Normal);
                case object o:
                    PSObject ps = PSObject.AsPSObject(o);
                    if (ps.Properties["FamilyName"] != null)
                    {
                        SKFontStyleWeight weight = ps.Properties["FontWeight"] == null ?
                            SKFontStyleWeight.Normal : LanguagePrimitives.ConvertTo<SKFontStyleWeight>(
                                ps.Properties["FontWeight"]);
                        SKFontStyleSlant slant = ps.Properties["FontSlant"] == null ?
                            SKFontStyleSlant.Upright : LanguagePrimitives.ConvertTo<SKFontStyleSlant>(
                                ps.Properties["FontSlant"]);
                        SKFontStyleWidth width = ps.Properties["FontWidth"] == null ?
                            SKFontStyleWidth.Normal : LanguagePrimitives.ConvertTo<SKFontStyleWidth>(
                                ps.Properties["FontWidth"]);
                        string familyName = LanguagePrimitives.ConvertTo<string>(ps.Properties["FamilyName"]);

                        return WordCloudUtils.FontManager.MatchFamily(familyName, new SKFontStyle(weight, width, slant));
                    }

                    break;
            }

            throw new ArgumentTransformationMetadataException();
        }
    }
}
