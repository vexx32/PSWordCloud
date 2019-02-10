using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Management.Automation;
using System.Reflection;
using System.Runtime.CompilerServices;
using SkiaSharp;

[assembly: InternalsVisibleTo("Cmdlet.Tests")]
namespace PSWordCloud
{
    internal enum WordOrientation : sbyte
    {
        Horizontal,
        Vertical,
        VerticalFlipped,
    }

    internal static class WCUtils
    {
        public static float ToRadians(this float degrees)
        {
            return (float)(degrees * Math.PI / 180);
        }

        public static void Shuffle<T>(this Random rng, T[] array)
        {
            int n = array.Length;
            while (n > 1)
            {
                int k = rng.Next(n--);
                T temp = array[n];
                array[n] = array[k];
                array[k] = temp;
            }
        }

        /// <summary>
        /// Checks if any part of the rectangle lies outside the region's bounds.
        /// </summary>
        /// <param name="region">The region to test for edge intersection.</param>
        /// <param name="other">The rectangle to test position against the edges of the region.</param>
        /// <returns>Returns false if the rectangle is entirely within the region, and false otherwise.</returns>
        public static bool FallsOutside(this SKRect other, SKRegion region)
        {
            var bounds = region.Bounds;
            return other.Top < bounds.Top
                || other.Bottom > bounds.Bottom
                || other.Left < bounds.Left
                || other.Right > bounds.Right;
        }

        public static void NextWord(this SKPaint brush, float wordSize, float strokeWidth)
            => brush.NextWord(wordSize, strokeWidth, SKColors.Black);

        public static void NextWord(this SKPaint brush, float wordSize, float strokeWidth, SKColor color)
        {
            brush.TextSize = wordSize;
            brush.StrokeWidth = wordSize * strokeWidth / 50;
            brush.IsStroke = false;
            brush.IsVerticalText = false;
            brush.Color = color;
        }

        public static float SortValue(this SKColor color, float sortAdjustment)
        {
            color.ToHsv(out float h, out float saturation, out float brightness);
            var rand = brightness * (sortAdjustment - 0.5f) / (1 - saturation);
            return brightness + rand;
        }

        public static bool SetPath(this SKRegion region, SKPath path, bool usePathBounds)
        {
            if (usePathBounds && path.GetBounds(out SKRect bounds))
            {
                using (SKRegion clip = new SKRegion())
                {
                    clip.SetRect(SKRectI.Ceiling(bounds));
                    return region.SetPath(path, clip);
                }
            }
            else
            {
                return region.SetPath(path);
            }
        }

        public static bool Op(this SKRegion region, SKPath path, SKRegionOperation operation)
        {
            using (SKRegion pathRegion = new SKRegion())
            {
                pathRegion.SetPath(path, true);
                return region.Op(pathRegion, operation);
            }
        }

        public static bool IntersectsRect(this SKRegion region, SKRect rect)
        {
            if (region.Bounds.IsEmpty)
            {
                return false;
            }

            using (SKRegion rectRegion = new SKRegion())
            {
                rectRegion.SetRect(SKRectI.Round(rect));
                return region.Intersects(rectRegion);
            }
        }

        public static bool IntersectsPath(this SKRegion region, SKPath path)
        {
            if (region.Bounds.IsEmpty)
            {
                return false;
            }

            using (SKRegion pathRegion = new SKRegion())
            {
                pathRegion.SetPath(path, region);
                return region.Intersects(pathRegion);
            }
        }

        internal static readonly ReadOnlyDictionary<string, SKColor> ColorLibrary =
            new ReadOnlyDictionary<string, SKColor>(typeof(SKColors)
            .GetFields(BindingFlags.Static | BindingFlags.Public)
            .ToDictionary(
                (field => field.Name),
                (field => (SKColor)field.GetValue(null)),
                StringComparer.OrdinalIgnoreCase));

        internal static readonly IEnumerable<string> ColorNames = ColorLibrary.Keys;

        internal static readonly IEnumerable<SKColor> StandardColors = ColorLibrary.Values;

        internal static SKColor GetColorByName(string colorName)
        {
            return ColorLibrary[colorName];
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

        public static SKFontManager FontManager = SKFontManager.Default;

        internal static ReadOnlyDictionary<string, (string Tooltip, SKSizeI Size)> StandardImageSizes =
            new ReadOnlyDictionary<string, (string, SKSizeI)>(new Dictionary<string, (string, SKSizeI)>() {
                {"480x800",     ("Mobile Screen Size (small)",  new SKSizeI(480, 800)  )},
                {"640x1146",    ("Mobile Screen Size (medium)", new SKSizeI(640, 1146) )},
                {"720p",        ("Standard HD 1280x720",        new SKSizeI(1280, 720) )},
                {"1080p",       ("Full HD 1920x1080",           new SKSizeI(1920, 1080))},
                {"4K",          ("Ultra HD 3840x2160",          new SKSizeI(3840, 2160))},
                {"A4",          ("816x1056",                    new SKSizeI(816, 1056) )},
                {"Poster11x17", ("1056x1632",                   new SKSizeI(1056, 1632))},
                {"Poster18x24", ("1728x2304",                   new SKSizeI(1728, 2304))},
                {"Poster24x36", ("2304x3456",                   new SKSizeI(2304, 3456))},
            });
    }
}
