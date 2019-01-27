using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Management.Automation;
using System.Reflection;
using SkiaSharp;

namespace PSWordCloud
{
    internal enum WordOrientation : sbyte
    {
        Horizontal,
        Vertical
    }

    internal enum ScanDirection : sbyte
    {
        ClockWise = 1,
        CounterClockwise = -1
    }
    static class WCUtils
    {
        public static float ToRadians(this float degrees)
        {
            return (float)(degrees * Math.PI / 180);
        }

        internal static readonly ReadOnlyDictionary<string, SKColor> ColorLibrary =
            new ReadOnlyDictionary<string, SKColor>(typeof(SKColors)
            .GetFields(BindingFlags.Static | BindingFlags.Public)
            .ToDictionary((field => field.Name), (field => (SKColor)field.GetValue(null))));

        internal static readonly IEnumerable<string> ColorNames = ColorLibrary.Keys;

        internal static readonly IEnumerable<SKColor> StandardColors = ColorLibrary.Values;

        internal static SKColor GetColorByName(string colorName)
        {
            return ColorLibrary[colorName];
        }

        internal static SKRectI ToSKRectI(this SKRect rectangle)
        {
            return new SKRectI(
                (int)Math.Round(rectangle.Left),
                (int)Math.Round(rectangle.Top),
                (int)Math.Round(rectangle.Right),
                (int)Math.Round(rectangle.Bottom));
        }

        internal static object GetValue(this IEnumerable collection, string key)
        {
            switch (collection)
            {
                case PSMemberInfoCollection<PSPropertyInfo> properties:
                    return properties[key].Value;
                case IDictionary dictionary:
                    return dictionary[key];
                case IDictionary<string, dynamic> dictT:
                    return dictT[key];
                default:
                    throw new ArgumentException(
                        string.Format(
                            "GetValue method only accepts {0} or {1}",
                            typeof(PSMemberInfoCollection<PSPropertyInfo>).ToString(),
                            typeof(IDictionary).ToString()));
            }
        }

        internal static SKFontManager FontManager = SKFontManager.Default;

        internal static ReadOnlyDictionary<string, (string Tooltip, SKSize Size)> StandardImageSizes =
            new ReadOnlyDictionary<string, (string, SKSize)>(new Dictionary<string, (string, SKSize)>() {
                {"480x800",     ("Mobile Screen Size (small)",  new SKSize(480, 800)  )},
                {"640x1146",    ("Mobile Screen Size (medium)", new SKSize(640, 1146) )},
                {"720p",        ("Standard HD 1280x720",        new SKSize(1280, 720) )},
                {"1080p",       ("Full HD 1920x1080",           new SKSize(1920, 1080))},
                {"4K",          ("Ultra HD 3840x2160",          new SKSize(3840, 2160))},
                {"A4",          ("816x1056",                    new SKSize(816, 1056) )},
                {"Poster11x17", ("1056x1632",                   new SKSize(1056, 1632))},
                {"Poster18x24", ("1728x2304",                   new SKSize(1728, 2304))},
                {"Poster24x36", ("2304x3456",                   new SKSize(2304, 3456))},
            });
    }
}
