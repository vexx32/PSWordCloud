using System;
using System.Collections;
using System.Management.Automation;
using System.Text.RegularExpressions;
using SkiaSharp;

namespace PSWordCloud.Attributes
{
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
            if (Utils.StandardImageSizes.ContainsKey(str))
            {
                return Utils.StandardImageSizes[str].Size;
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
}
