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
            foreach (var result in WCUtils.StandardImageSizes)
            {
                if (string.IsNullOrEmpty(wordToComplete) ||
                    result.Key.StartsWith(wordToComplete, StringComparison.OrdinalIgnoreCase))
                {
                    yield return new CompletionResult(
                        result.Key,
                        result.Key,
                        CompletionResultType.ParameterValue,
                        result.Value.Tooltip);
                }
            }
        }
    }

    public class TransformToSKSizeIAttribute : ArgumentTransformationAttribute
    {
        public override object Transform(EngineIntrinsics engineIntrinsics, object inputData)
        {
            int sideLength = 0;
            SKSizeI? result;

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
                    result = GetSizeFromString(s);
                    if (result is not null)
                    {
                        return result;
                    }

                    break;
                case object o:
                    result = GetSizeFromProperties(o);
                    if (result is not null)
                    {
                        return result;
                    }

                    break;
            }

            if (sideLength > 0)
            {
                return new SKSizeI(sideLength, sideLength);
            }

            var errorMessage = $"Unrecognisable input '{inputData}' for SKSize parameter. See the help documentation for the parameter for allowed values.";
            throw new ArgumentTransformationMetadataException(errorMessage);
        }

        private SKSizeI? GetSizeFromString(string str)
        {
            if (WCUtils.StandardImageSizes.ContainsKey(str))
            {
                return WCUtils.StandardImageSizes[str].Size;
            }
            else
            {
                var matchWH = Regex.Match(str, @"^(?<Width>[\d\.,]+)x(?<Height>[\d\.,]+)(px)?$");
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
                            "Could not parse input string as an integer value", e);
                    }
                }

                var matchSide = Regex.Match(str, @"^(?<SideLength>[\d\.,]+)(px)?$");
                if (matchSide.Success)
                {
                    var sideLength = int.Parse(matchSide.Groups["SideLength"].Value);
                    return new SKSizeI(sideLength, sideLength);
                }
            }

            return null;
        }

        private SKSizeI? GetSizeFromProperties(object obj)
        {
            IEnumerable properties;
            if (obj is Hashtable ht)
            {
                properties = ht;
            }
            else
            {
                properties = PSObject.AsPSObject(obj).Properties;
            }

            if (properties.GetValue("Width") is not null && properties.GetValue("Height") is not null)
            {
                // If these conversions fail, the exception will cause the transform to fail.
                object? width = properties.GetValue("Width");
                object? height = properties.GetValue("Height");

                if (width is null || height is null)
                {
                    return null;
                }

                return new SKSizeI(width.ConvertTo<int>(), height.ConvertTo<int>());
            }

            return null;
        }
    }

    public class FontFamilyCompleter : IArgumentCompleter
    {
        public IEnumerable<CompletionResult> CompleteArgument(
            string commandName,
            string parameterName,
            string wordToComplete,
            CommandAst commandAst,
            IDictionary fakeBoundParameters)
        {
            string matchString = wordToComplete.TrimStart('"').TrimEnd('"');
            foreach (string font in WCUtils.FontList)
            {
                if (string.IsNullOrEmpty(wordToComplete)
                    || font.StartsWith(matchString, StringComparison.OrdinalIgnoreCase))
                {
                    if (font.Contains(' ') || font.Contains('#') || wordToComplete.StartsWith("\""))
                    {
                        var result = string.Format("\"{0}\"", font);
                        yield return new CompletionResult(result, font, CompletionResultType.ParameterName, font);
                    }
                    else
                    {
                        yield return new CompletionResult(font, font, CompletionResultType.ParameterName, font);
                    }
                }
            }
        }
    }

    public class TransformToSKTypefaceAttribute : ArgumentTransformationAttribute
    {
        public override object Transform(EngineIntrinsics engineIntrinsics, object inputData)
        {
            return inputData switch
            {
                SKTypeface t => t,
                string s => WCUtils.FontManager.MatchFamily(s, SKFontStyle.Normal),
                _ => CreateTypefaceFromObject(inputData),
            };
        }

        private static SKTypeface CreateTypefaceFromObject(object input)
        {
            IEnumerable properties;
            if (input is Hashtable ht)
            {
                properties = ht;
            }
            else
            {
                properties = PSObject.AsPSObject(input).Properties;
            }

            SKFontStyle style;
            if (properties.GetValue("FontWeight") is not null
                || properties.GetValue("FontSlant") is not null
                || properties.GetValue("FontWidth") is not null)
            {
                object? weightValue = properties.GetValue("FontWeight");
                SKFontStyleWeight weight = weightValue is null
                    ? SKFontStyleWeight.Normal
                    : weightValue.ConvertTo<SKFontStyleWeight>();

                object? slantValue = properties.GetValue("FontSlant");
                SKFontStyleSlant slant = slantValue is null
                    ? SKFontStyleSlant.Upright
                    : slantValue.ConvertTo<SKFontStyleSlant>();

                object? widthValue = properties.GetValue("FontWidth");
                SKFontStyleWidth width = widthValue is null
                    ? SKFontStyleWidth.Normal
                    : widthValue.ConvertTo<SKFontStyleWidth>();

                style = new SKFontStyle(weight, width, slant);
            }
            else
            {
                style = properties.GetValue("FontStyle") is SKFontStyle customStyle
                    ? customStyle
                    : SKFontStyle.Normal;
            }

            string familyName = properties.GetValue("FamilyName")?.ConvertTo<string>() ?? string.Empty;
            return WCUtils.FontManager.MatchFamily(familyName, style);
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
                if (string.IsNullOrEmpty(wordToComplete)
                    || color.StartsWith(wordToComplete, StringComparison.OrdinalIgnoreCase))
                {
                    SKColor colorValue = WCUtils.ColorLibrary[color];
                    yield return new CompletionResult(
                        completionText: color,
                        listItemText: color,
                        CompletionResultType.ParameterValue,
                        toolTip: string.Format(
                            "{0} (R: {1}, G: {2}, B: {3}, A: {4})",
                            color, colorValue.Red, colorValue.Green, colorValue.Blue, colorValue.Alpha));
                }
            }
        }
    }

    public class TransformToSKColorAttribute : ArgumentTransformationAttribute
    {
        private SKColor[] TransformObject(object input)
        {
            var colorList = new List<SKColor>();
            object[] array;
            if (input is object[] o)
            {
                array = o;
            }
            else
            {
                array = new[] { input };
            }

            foreach (object item in array)
            {
                if (item is string str)
                {
                    colorList.AddRange(MatchColor(str));
                    continue;
                }

                IEnumerable properties;
                if (item is Hashtable ht)
                {
                    properties = ht;
                }
                else
                {
                    properties = PSObject.AsPSObject(item).Properties;
                }

                object? redValue = properties.GetValue("red");
                byte red = redValue is null
                    ? (byte)0
                    : redValue.ConvertTo<byte>();

                object? greenValue = properties.GetValue("green");
                byte green = greenValue is null
                    ? (byte)0
                    : greenValue.ConvertTo<byte>();

                object? blueValue = properties.GetValue("blue");
                byte blue = blueValue is null
                    ? (byte)0
                    : blueValue.ConvertTo<byte>();

                object? alphaValue = properties.GetValue("alpha");
                byte alpha = alphaValue is null
                    ? (byte)255
                    : alphaValue.ConvertTo<byte>();

                colorList.Add(new SKColor(red, green, blue, alpha));
            }

            return colorList.ToArray();
        }

        private SKColor[] MatchColor(string name)
        {
            string errorMessage;
            if (WCUtils.ColorNames.Contains(name))
            {
                return new[] { WCUtils.ColorLibrary[name] };
            }

            if (string.Equals(name, "transparent", StringComparison.OrdinalIgnoreCase))
            {
                return new[] { SKColor.Empty };
            }

            if (WildcardPattern.ContainsWildcardCharacters(name))
            {
                var pattern = new WildcardPattern(name, WildcardOptions.IgnoreCase);

                var colorList = new List<SKColor>();
                foreach (var color in WCUtils.ColorLibrary)
                {
                    if (pattern.IsMatch(color.Key))
                    {
                        colorList.Add(color.Value);
                    }
                }

                if (colorList.Count > 0)
                {
                    return colorList.ToArray();
                }

                errorMessage = $"Wildcard pattern '{name}' did not match any known color names.";
                throw new ArgumentTransformationMetadataException(errorMessage);
            }

            if (SKColor.TryParse(name, out SKColor c))
            {
                return new[] { c };
            }

            errorMessage = $"Unrecognised color name: '{name}'.";
            throw new ArgumentTransformationMetadataException(errorMessage);
        }

        private object Normalize(SKColor[] results)
        {
            if (results.Length == 1)
            {
                return results[0];
            }
            else
            {
                return results;
            }
        }

        public override object Transform(EngineIntrinsics engineIntrinsics, object inputData)
        {
            return inputData switch
            {
                string s => Normalize(MatchColor(s)),
                SKColor color => color,
                _ => Normalize(TransformObject(inputData)),
            };
        }
    }

    public class AngleCompleter : IArgumentCompleter
    {
        public IEnumerable<CompletionResult> CompleteArgument(
            string commandName,
            string parameterName,
            string wordToComplete,
            CommandAst commandAst,
            IDictionary fakeBoundParameters)
        {
            for (float angle = 0; angle <= 360; angle += 45)
            {
                var s = angle.ConvertTo<string>();
                yield return new CompletionResult(s, s, CompletionResultType.ParameterValue, s);
            }
        }
    }
}
