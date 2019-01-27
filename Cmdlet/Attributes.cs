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
    internal class ImageSizeCompleter : IArgumentCompleter
    {
        public IEnumerable<CompletionResult> CompleteArgument(
            string commandName,
            string parameterName,
            string wordToComplete,
            CommandAst commandAst,
            IDictionary fakeboundParameters)
        {

            var matchingResults = WCUtils.StandardImageSizes.Where(
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

    public class TransformToSKSizeIAttribute : ArgumentTransformationAttribute
    {
        public override object Transform(EngineIntrinsics engineIntrinsics, object inputData)
        {
            int sideLength = 0;
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
                    if (u <= int.MaxValue)
                    {
                        sideLength = (int)u;
                    }

                    break;
                case long l:
                    if (l <= int.MaxValue)
                    {
                        sideLength = (int)l;
                    }

                    break;
                case ulong ul:
                    if (ul <= int.MaxValue)
                    {
                        sideLength = (int)ul;
                    }
                    break;
                case decimal d:
                    if (d <= int.MaxValue)
                    {
                        sideLength = (int)Math.Round(d);
                    }
                    break;
                case float f:
                    if (f <= int.MaxValue)
                    {
                        sideLength = (int)Math.Round(f);
                    }
                    break;
                case double d:
                    if (d <= int.MaxValue)
                    {
                        sideLength = (int)Math.Round(d);
                    }
                    break;
                case string s:
                    if (WCUtils.StandardImageSizes.ContainsKey(s))
                    {
                        return WCUtils.StandardImageSizes[s].Size;
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
                    IEnumerable properties = null;
                    if (o is Hashtable ht)
                    {
                        properties = ht;
                    }
                    else
                    {
                        properties = PSObject.AsPSObject(o).Properties;
                    }

                    if (properties.GetValue("Width") != null && properties.GetValue("Height") != null)
                    {
                        // If these conversions fail, the exception will cause the transform to fail.
                        var width = LanguagePrimitives.ConvertTo<int>(properties.GetValue("Width"));
                        var height = LanguagePrimitives.ConvertTo<int>(properties.GetValue("Height"));

                        return new SKSizeI(width, height);
                    }

                    break;
            }

            if (sideLength > 0)
            {
                return new SKSizeI(sideLength, sideLength);
            }

            throw new ArgumentTransformationMetadataException();
        }
    }

    internal class FontFamilyCompleter : IArgumentCompleter
    {
        private static List<string> _fontList = new List<string>(WCUtils.FontManager.GetFontFamilies());

        public IEnumerable<CompletionResult> CompleteArgument(
            string commandName,
            string parameterName,
            string wordToComplete,
            CommandAst commandAst,
            IDictionary fakeBoundParameters)
        {
            var fonts = WCUtils.FontManager.GetFontFamilies();
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

    public class TransformToSKTypefaceAttribute : ArgumentTransformationAttribute
    {
        public override object Transform(EngineIntrinsics engineIntrinsics, object inputData)
        {
            switch (inputData)
            {
                case SKTypeface t:
                    return t;
                case string s:
                    return WCUtils.FontManager.MatchFamily(s, SKFontStyle.Normal);
                default:
                    IEnumerable properties = null;
                    if (inputData is Hashtable ht)
                    {
                        properties = ht;
                    }
                    else
                    {
                        properties = PSObject.AsPSObject(inputData).Properties;
                    }

                    SKFontStyleWeight weight = properties.GetValue("FontWeight") == null ?
                        SKFontStyleWeight.Normal : LanguagePrimitives.ConvertTo<SKFontStyleWeight>(
                            properties.GetValue("FontWeight"));
                    SKFontStyleSlant slant = properties.GetValue("FontSlant") == null ?
                        SKFontStyleSlant.Upright : LanguagePrimitives.ConvertTo<SKFontStyleSlant>(
                            properties.GetValue("FontSlant"));
                    SKFontStyleWidth width = properties.GetValue("FontWidth") == null ?
                        SKFontStyleWidth.Normal : LanguagePrimitives.ConvertTo<SKFontStyleWidth>(
                            properties.GetValue("FontWidth"));
                    string familyName = LanguagePrimitives.ConvertTo<string>(properties.GetValue("FamilyName"));

                    return WCUtils.FontManager.MatchFamily(familyName, new SKFontStyle(weight, width, slant));
            }
        }
    }

    public class SKColorCompleter : IArgumentCompleter
    {
        public IEnumerable<CompletionResult> CompleteArgument(
            string commandName,
            string parameterName,
            string wordToComplete,
            CommandAst commandAst,
            IDictionary fakeBoundParameters)
        {
            foreach (string color in WCUtils.ColorNames)
            {
                SKColor colorValue = WCUtils.ColorLibrary[color];
                yield return new CompletionResult(
                    color,
                    color,
                    CompletionResultType.ParameterValue,
                    string.Format("{0} (R: {1}, G: {2}, B: {3}, A: {4})",
                        color, colorValue.Red, colorValue.Green, colorValue.Blue, colorValue.Alpha));
            }
        }
    }

    public class TransformToSKColorAttribute : ArgumentTransformationAttribute
    {
        public override object Transform(EngineIntrinsics engineIntrinsics, object inputData)
        {
            switch (inputData)
            {
                case string s:
                    if (WCUtils.ColorNames.Contains(s))
                    {
                        return WCUtils.ColorLibrary[s];
                    }

                    if (string.Equals(s, "transparent", StringComparison.OrdinalIgnoreCase))
                    {
                        return SKColor.Empty;
                    }

                    if (SKColor.TryParse(s, out SKColor c))
                    {
                        return c;
                    }

                    throw new ArgumentTransformationMetadataException();
                case SKColor color:
                    return color;
                default:
                    IEnumerable properties = null;
                    if (inputData is Hashtable ht)
                    {
                        properties = ht;
                    }
                    else
                    {
                        properties = PSObject.AsPSObject(inputData).Properties;
                    }

                    byte red = properties.GetValue("red") == null ?
                        (byte)0 : (byte)LanguagePrimitives.ConvertTo<byte>(properties.GetValue("red"));
                    byte green = properties.GetValue("green") == null ?
                        (byte)0 : (byte)LanguagePrimitives.ConvertTo<byte>(properties.GetValue("green"));
                    byte blue = properties.GetValue("blue") == null ?
                        (byte)0 : (byte)LanguagePrimitives.ConvertTo<byte>(properties.GetValue("blue"));
                    byte alpha = properties.GetValue("alpha") == null ?
                        (byte)0 : (byte)LanguagePrimitives.ConvertTo<byte>(properties.GetValue("alpha"));

                    return new SKColor(red, green, blue, alpha);
            }
        }
    }

}
