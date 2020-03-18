using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Management.Automation;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;
using System.Xml;
using SkiaSharp;

[assembly: InternalsVisibleTo("PSWordCloud.Tests")]
namespace PSWordCloud
{
    public enum WordOrientations : sbyte
    {
        None,
        Vertical,
        FlippedVertical,
        EitherVertical,
        UprightDiagonals,
        InvertedDiagonals,
        AllDiagonals,
        AllUpright,
        AllInverted,
        All,
    }

    public enum WordBubbleShape : sbyte
    {
        None,
        Rectangle,
        Square,
        Circle,
        Oval
    }

    internal static class WCUtils
    {
        internal static SKPoint Multiply(this SKPoint point, float factor)
            => new SKPoint(point.X * factor, point.Y * factor);

        internal static float ToRadians(this float degrees)
            => (float)(degrees * Math.PI / 180);

        /// <summary>
        /// Returns a font scale value based on the size of the letter X in a given typeface.
        /// </summary>
        /// <param name="typeface">The typeface to measure the scale from.</param>
        /// <returns>A float value typically between 0 and 1. Many common typefaces have values around 0.5.</returns>
        internal static float GetFontScale(SKTypeface typeface)
        {
            var text = "X";
            using var paint = new SKPaint
            {
                Typeface = typeface,
                TextSize = 1
            };
            var rect = paint.GetTextPath(text, 0, 0).ComputeTightBounds();

            return (rect.Width + rect.Height) / 2;
        }

        internal static bool IsDistinctColor(this SKColor target, SKColor backdrop)
        {
            backdrop.ToHsv(out float refHue, out float refSaturation, out float refBrightness);
            target.ToHsv(out float hue, out float saturation, out float brightness);

            float brightnessDistance = Math.Abs(refBrightness - brightness);
            if (brightnessDistance > 24)
            {
                return true;
            }

            if (Math.Abs(refHue - hue) > 18 && brightnessDistance > 24)
            {
                return true;
            }

            if (Math.Abs(refSaturation - saturation) > 24 && brightnessDistance > 12)
            {
                return true;
            }

            if (target.Alpha == 0)
            {
                return false;
            }

            return false;
        }

        internal static void Shuffle<T>(this Random rng, T[] array)
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
        internal static bool FallsOutside(this SKRect other, SKRegion region)
        {
            var bounds = region.Bounds;
            return other.Top < bounds.Top
                || other.Bottom > bounds.Bottom
                || other.Left < bounds.Left
                || other.Right > bounds.Right;
        }

        internal static void NextWord(this SKPaint brush, float wordSize, float strokeWidth)
            => brush.NextWord(wordSize, strokeWidth, SKColors.Black);

        internal static void NextWord(this SKPaint brush, float wordSize, float strokeWidth, SKColor color)
        {
            brush.TextSize = wordSize;
            brush.IsStroke = false;
            brush.Style = SKPaintStyle.StrokeAndFill;
            brush.StrokeWidth = wordSize * strokeWidth * NewWordCloudCommand.STROKE_BASE_SCALE;
            brush.IsVerticalText = false;
            brush.Color = color;
        }

        internal static float SortValue(this SKColor color, float sortAdjustment)
        {
            color.ToHsv(out _, out float saturation, out float brightness);
            var rand = brightness * (sortAdjustment - 0.5f) / (1 - saturation);

            return brightness + rand;
        }

        internal static bool SetPath(this SKRegion region, SKPath path, bool usePathBounds)
        {
            if (usePathBounds && path.GetBounds(out SKRect bounds))
            {
                using SKRegion clip = new SKRegion();

                clip.SetRect(SKRectI.Ceiling(bounds));
                return region.SetPath(path, clip);
            }
            else
            {
                return region.SetPath(path);
            }
        }

        internal static bool Op(this SKRegion region, SKPath path, SKRegionOperation operation)
        {
            using SKRegion pathRegion = new SKRegion();

            pathRegion.SetPath(path, true);
            return region.Op(pathRegion, operation);
        }

        internal static bool IntersectsRect(this SKRegion region, SKRect rect)
        {
            if (region.Bounds.IsEmpty)
            {
                return false;
            }

            using SKRegion rectRegion = new SKRegion();

            rectRegion.SetRect(SKRectI.Round(rect));
            return region.Intersects(rectRegion);
        }

        internal static bool IntersectsPath(this SKRegion region, SKPath path)
        {
            if (region.Bounds.IsEmpty)
            {
                return false;
            }

            using SKRegion pathRegion = new SKRegion();

            pathRegion.SetPath(path, region);
            return region.Intersects(pathRegion);
        }

        internal static SKRect GetEnclosingSquare(this SKRect rect)
        {
            // Inflate the smaller dimension
            if (rect.Width > rect.Height)
            {
                return SKRect.Inflate(rect, x: 0, y: (rect.Width - rect.Height) / 2);
            }

            if (rect.Height > rect.Width)
            {
                return SKRect.Inflate(rect, x: (rect.Height - rect.Width) / 2, y: 0);
            }

            // It was already a square, but we need to return a copy
            return SKRect.Create(rect.Location, rect.Size);
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

        internal static string GetPrettyString(this XmlDocument document)
        {
            var stringBuilder = new StringBuilder();

            var settings = new XmlWriterSettings
            {
                Indent = true
            };

            using (var xmlWriter = XmlWriter.Create(stringBuilder, settings))
            {
                document.Save(xmlWriter);
            }

            return stringBuilder.ToString();
        }

        internal static object GetValue(this IEnumerable collection, string key)
        {
            return collection switch
            {
                PSMemberInfoCollection<PSPropertyInfo> properties => properties[key].Value,
                IDictionary dictionary => dictionary[key],
                IDictionary<string, dynamic> dictT => dictT[key],
                _ => throw new ArgumentException(
                    string.Format(
                    "GetValue method only accepts {0} or {1}",
                    typeof(PSMemberInfoCollection<PSPropertyInfo>).ToString(),
                    typeof(IDictionary).ToString())),
            };
        }

        internal static SKFontManager FontManager = SKFontManager.Default;

        internal static IEnumerable<string> FontList = WCUtils.FontManager.FontFamilies
            .OrderBy(name => name, StringComparer.OrdinalIgnoreCase);

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
