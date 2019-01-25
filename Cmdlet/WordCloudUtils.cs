using System;
using System.Collections.Generic;
using SkiaSharp;

namespace PSWordCloud
{
    static class WordCloudUtils
    {
        public static SKRectI ToSKRectI(this SKRect rectangle)
        {
            return new SKRectI((int)rectangle.Left, (int)rectangle.Top, (int)rectangle.Right, (int)rectangle.Bottom);
        }

        public static SKFontManager FontManager = SKFontManager.Default;

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
    }
}
