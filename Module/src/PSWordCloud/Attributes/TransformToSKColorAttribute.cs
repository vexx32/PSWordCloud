using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using SkiaSharp;

namespace PSWordCloud.Attributes
{
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
            if (Utils.ColorNames.Contains(name))
            {
                return new[] { Utils.ColorLibrary[name] };
            }

            if (string.Equals(name, "transparent", StringComparison.OrdinalIgnoreCase))
            {
                return new[] { SKColor.Empty };
            }

            if (WildcardPattern.ContainsWildcardCharacters(name))
            {
                var pattern = new WildcardPattern(name, WildcardOptions.IgnoreCase);

                var colorList = new List<SKColor>();
                foreach (var color in Utils.ColorLibrary)
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
}
