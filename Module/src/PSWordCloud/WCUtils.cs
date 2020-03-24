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

        /// <summary>
        /// Determines whether a given color is considered sufficiently visually distinct from a backdrop color.
        /// </summary>
        /// <param name="target">The target color.</param>
        /// <param name="backdrop">A reference color to compare against.</param>
        internal static bool IsDistinctColor(this SKColor target, SKColor backdrop)
        {
            backdrop.ToHsv(out float refHue, out float refSaturation, out float refBrightness);
            target.ToHsv(out float hue, out float saturation, out float brightness);

            float brightnessDistance = Math.Abs(refBrightness - brightness);
            if (brightnessDistance > 30)
            {
                return true;
            }

            if (Math.Abs(refHue - hue) > 24 && brightnessDistance > 20)
            {
                return true;
            }

            if (Math.Abs(refSaturation - saturation) > 24 && brightnessDistance > 18)
            {
                return true;
            }

            if (target.Alpha == 0)
            {
                return false;
            }

            return false;
        }

        /// <summary>
        /// Performs an in-place random shuffle on an array by swapping elements.
        /// This algorithm is pretty commonly used, but effective and fast enough for our purposes.
        /// </summary>
        /// <param name="rng">Random number generator.</param>
        /// <param name="array">The array to shuffle.</param>
        /// <typeparam name="T">The element type of the array.</typeparam>
        internal static T[] Shuffle<T>(this Random rng, T[] array)
        {
            int n = array.Length;
            while (n > 1)
            {
                int k = rng.Next(n--);
                T temp = array[n];
                array[n] = array[k];
                array[k] = temp;
            }

            return array;
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

        /// <summary>
        /// Checks if the given <paramref name="point"/> lies somewhere inside the <paramref name="region"/>.
        /// </summary>
        /// <param name="region">The region that defines the bounds.</param>
        /// <param name="point">The point to check.</param>
        /// <returns></returns>
        internal static bool Contains(this SKRegion region, SKPoint point)
        {
            SKRectI bounds = region.Bounds;
            return bounds.Left < point.X && point.X < bounds.Right
                && bounds.Top < point.Y && point.Y < bounds.Bottom;
        }

        internal static float GetEnclosedArea(this SKPath path)
        {
            if (path.IsEmpty)
            {
                return 0;
            }

            SKRect bounds = path.TightBounds;
            var boundedArea = bounds.Width * bounds.Height;
            var totalPoints = 10000;
            var enclosedPoints = 0;

            for (float x = bounds.Left; x < bounds.Right; x += bounds.Width / (float)Math.Sqrt(totalPoints))
            {
                for (float y = bounds.Top; y < bounds.Bottom; y += bounds.Height / (float)Math.Sqrt(totalPoints))
                {
                    if (path.Contains(x, y))
                    {
                        enclosedPoints++;
                    }
                }
            }

            var enclosedAreaRatio = enclosedPoints / totalPoints;

            return enclosedAreaRatio * boundedArea;
        }

        /// <summary>
        /// Prepares the brush to draw the next word.
        /// This overload assumes the text to be drawn will be black.
        /// </summary>
        /// <param name="brush"></param>
        /// <param name="wordSize"></param>
        /// <param name="strokeWidth"></param>
        internal static void Prepare(this SKPaint brush, float wordSize, float strokeWidth)
            => brush.Prepare(wordSize, strokeWidth, SKColors.Black);

        /// <summary>
        /// Prepares the brush to draw the next word.
        /// </summary>
        /// <param name="brush">The brush to prepare.</param>
        /// <param name="wordSize">The size of the word we'll be drawing.</param>
        /// <param name="strokeWidth">Width of the stroke we'll be drawing.</param>
        /// <param name="color">Color of the word we'll be drawing.</param>
        internal static void Prepare(this SKPaint brush, float wordSize, float strokeWidth, SKColor color)
        {
            brush.TextSize = wordSize;
            brush.IsStroke = false;
            brush.Style = SKPaintStyle.StrokeAndFill;
            brush.StrokeWidth = wordSize * strokeWidth * NewWordCloudCommand.STROKE_BASE_SCALE;
            brush.IsVerticalText = false;
            brush.Color = color;
        }

        /// <summary>
        /// Sets the contents of the region to the specified path.
        /// </summary>
        /// <param name="region">The region to set the path into.</param>
        /// <param name="path">The path object.</param>
        /// <param name="usePathBounds">Whether to set the region's new bounds to the bounds of the path itself.</param>
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

        /// <summary>
        /// Combines the region with a given path, specifying the operation used to combine.
        /// </summary>
        /// <param name="region">The region to perform the operation on.</param>
        /// <param name="path">The path to perform the operation with.</param>
        /// <param name="operation">The type of operation to perform.</param>
        /// <returns></returns>
        internal static bool CombineWithPath(this SKRegion region, SKPath path, SKRegionOperation operation)
        {
            using SKRegion pathRegion = new SKRegion();

            pathRegion.SetPath(path, usePathBounds: true);
            return region.Op(pathRegion, operation);
        }

        /// <summary>
        /// Checks whether the region intersects the given rectangle.
        /// </summary>
        /// <param name="region">The region to check collision with.</param>
        /// <param name="rect">The rectangle to check for intersection.</param>
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

        /// <summary>
        /// Checks whether the region intersects the given path.
        /// </summary>
        /// <param name="region">The region to check collision with.</param>
        /// <param name="rect">The rectangle to check for intersection.</param>
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

        /// <summary>
        /// Gets the smallest square that would completely contain the given rectangle, with the rectangle positioned
        /// at its centre.
        /// </summary>
        /// <param name="rect">The rectangle to find the containing square for.</param>
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

        /// <summary>
        /// A list of standard color names supported for tab completion.
        /// </summary>
        internal static IEnumerable<string> ColorNames { get => ColorLibrary.Keys; }

        /// <summary>
        /// The complete set of standard colors.
        /// </summary>
        internal static IEnumerable<SKColor> StandardColors { get => ColorLibrary.Values; }

        /// <summary>
        /// Gets a color from the library by name.
        /// </summary>
        /// <param name="colorName"></param>
        /// <returns></returns>
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


        private static readonly Dictionary<string, SKColor> _library;
        /// <summary>
        /// A library mapping color names to SKColor instances.
        /// </summary>
        /// <remarks>The library contains both SkiaSharp.SKColors named colors and X11 named colors.</remarks>
        internal static ReadOnlyDictionary<string, SKColor> ColorLibrary { get => new ReadOnlyDictionary<string, SKColor>(_library); }

        static WCUtils()
        {
            _library = new Dictionary<string, SKColor>(StringComparer.CurrentCultureIgnoreCase)
            {
                { "aliceblue", SKColor.Parse("f0f8ff") },
                { "antiquewhite", SKColor.Parse("faebd7") },
                { "antiquewhite1", SKColor.Parse("ffefdb") },
                { "antiquewhite2", SKColor.Parse("eedfcc") },
                { "antiquewhite3", SKColor.Parse("cdc0b0") },
                { "antiquewhite4", SKColor.Parse("8b8378") },
                { "aquamarine1", SKColor.Parse("7fffd4") },
                { "aquamarine2", SKColor.Parse("76eec6") },
                { "aquamarine4", SKColor.Parse("458b74") },
                { "azure1", SKColor.Parse("f0ffff") },
                { "azure2", SKColor.Parse("e0eeee") },
                { "azure3", SKColor.Parse("c1cdcd") },
                { "azure4", SKColor.Parse("838b8b") },
                { "beige", SKColor.Parse("f5f5dc") },
                { "bisque1", SKColor.Parse("ffe4c4") },
                { "bisque2", SKColor.Parse("eed5b7") },
                { "bisque3", SKColor.Parse("cdb79e") },
                { "bisque4", SKColor.Parse("8b7d6b") },
                { "black", SKColor.Parse("000000") },
                { "blanchedalmond", SKColor.Parse("ffebcd") },
                { "blue1", SKColor.Parse("0000ff") },
                { "blue2", SKColor.Parse("0000ee") },
                { "blue4", SKColor.Parse("00008b") },
                { "blueviolet", SKColor.Parse("8a2be2") },
                { "brown", SKColor.Parse("a52a2a") },
                { "brown1", SKColor.Parse("ff4040") },
                { "brown2", SKColor.Parse("ee3b3b") },
                { "brown3", SKColor.Parse("cd3333") },
                { "brown4", SKColor.Parse("8b2323") },
                { "burlywood", SKColor.Parse("deb887") },
                { "burlywood1", SKColor.Parse("ffd39b") },
                { "burlywood2", SKColor.Parse("eec591") },
                { "burlywood3", SKColor.Parse("cdaa7d") },
                { "burlywood4", SKColor.Parse("8b7355") },
                { "cadetblue", SKColor.Parse("5f9ea0") },
                { "cadetblue1", SKColor.Parse("98f5ff") },
                { "cadetblue2", SKColor.Parse("8ee5ee") },
                { "cadetblue3", SKColor.Parse("7ac5cd") },
                { "cadetblue4", SKColor.Parse("53868b") },
                { "chartreuse1", SKColor.Parse("7fff00") },
                { "chartreuse2", SKColor.Parse("76ee00") },
                { "chartreuse3", SKColor.Parse("66cd00") },
                { "chartreuse4", SKColor.Parse("458b00") },
                { "chocolate", SKColor.Parse("d2691e") },
                { "chocolate1", SKColor.Parse("ff7f24") },
                { "chocolate2", SKColor.Parse("ee7621") },
                { "chocolate3", SKColor.Parse("cd661d") },
                { "coral", SKColor.Parse("ff7f50") },
                { "coral1", SKColor.Parse("ff7256") },
                { "coral2", SKColor.Parse("ee6a50") },
                { "coral3", SKColor.Parse("cd5b45") },
                { "coral4", SKColor.Parse("8b3e2f") },
                { "cornflowerblue", SKColor.Parse("6495ed") },
                { "cornsilk1", SKColor.Parse("fff8dc") },
                { "cornsilk2", SKColor.Parse("eee8cd") },
                { "cornsilk3", SKColor.Parse("cdc8b1") },
                { "cornsilk4", SKColor.Parse("8b8878") },
                { "cyan1", SKColor.Parse("00ffff") },
                { "cyan2", SKColor.Parse("00eeee") },
                { "cyan3", SKColor.Parse("00cdcd") },
                { "cyan4", SKColor.Parse("008b8b") },
                { "darkgoldenrod", SKColor.Parse("b8860b") },
                { "darkgoldenrod1", SKColor.Parse("ffb90f") },
                { "darkgoldenrod2", SKColor.Parse("eead0e") },
                { "darkgoldenrod3", SKColor.Parse("cd950c") },
                { "darkgoldenrod4", SKColor.Parse("8b6508") },
                { "darkgreen", SKColor.Parse("006400") },
                { "darkkhaki", SKColor.Parse("bdb76b") },
                { "darkolivegreen", SKColor.Parse("556b2f") },
                { "darkolivegreen1", SKColor.Parse("caff70") },
                { "darkolivegreen2", SKColor.Parse("bcee68") },
                { "darkolivegreen3", SKColor.Parse("a2cd5a") },
                { "darkolivegreen4", SKColor.Parse("6e8b3d") },
                { "darkorange", SKColor.Parse("ff8c00") },
                { "darkorange1", SKColor.Parse("ff7f00") },
                { "darkorange2", SKColor.Parse("ee7600") },
                { "darkorange3", SKColor.Parse("cd6600") },
                { "darkorange4", SKColor.Parse("8b4500") },
                { "darkorchid", SKColor.Parse("9932cc") },
                { "darkorchid1", SKColor.Parse("bf3eff") },
                { "darkorchid2", SKColor.Parse("b23aee") },
                { "darkorchid3", SKColor.Parse("9a32cd") },
                { "darkorchid4", SKColor.Parse("68228b") },
                { "darksalmon", SKColor.Parse("e9967a") },
                { "darkseagreen", SKColor.Parse("8fbc8f") },
                { "darkseagreen1", SKColor.Parse("c1ffc1") },
                { "darkseagreen2", SKColor.Parse("b4eeb4") },
                { "darkseagreen3", SKColor.Parse("9bcd9b") },
                { "darkseagreen4", SKColor.Parse("698b69") },
                { "darkslateblue", SKColor.Parse("483d8b") },
                { "darkslategray", SKColor.Parse("2f4f4f") },
                { "darkslategray1", SKColor.Parse("97ffff") },
                { "darkslategray2", SKColor.Parse("8deeee") },
                { "darkslategray3", SKColor.Parse("79cdcd") },
                { "darkslategray4", SKColor.Parse("528b8b") },
                { "darkturquoise", SKColor.Parse("00ced1") },
                { "darkviolet", SKColor.Parse("9400d3") },
                { "deeppink1", SKColor.Parse("ff1493") },
                { "deeppink2", SKColor.Parse("ee1289") },
                { "deeppink3", SKColor.Parse("cd1076") },
                { "deeppink4", SKColor.Parse("8b0a50") },
                { "deepskyblue1", SKColor.Parse("00bfff") },
                { "deepskyblue2", SKColor.Parse("00b2ee") },
                { "deepskyblue3", SKColor.Parse("009acd") },
                { "deepskyblue4", SKColor.Parse("00688b") },
                { "dimgray", SKColor.Parse("696969") },
                { "dodgerblue1", SKColor.Parse("1e90ff") },
                { "dodgerblue2", SKColor.Parse("1c86ee") },
                { "dodgerblue3", SKColor.Parse("1874cd") },
                { "dodgerblue4", SKColor.Parse("104e8b") },
                { "firebrick", SKColor.Parse("b22222") },
                { "firebrick1", SKColor.Parse("ff3030") },
                { "firebrick2", SKColor.Parse("ee2c2c") },
                { "firebrick3", SKColor.Parse("cd2626") },
                { "firebrick4", SKColor.Parse("8b1a1a") },
                { "floralwhite", SKColor.Parse("fffaf0") },
                { "forestgreen", SKColor.Parse("228b22") },
                { "gainsboro", SKColor.Parse("dcdcdc") },
                { "ghostwhite", SKColor.Parse("f8f8ff") },
                { "gold1", SKColor.Parse("ffd700") },
                { "gold2", SKColor.Parse("eec900") },
                { "gold3", SKColor.Parse("cdad00") },
                { "gold4", SKColor.Parse("8b7500") },
                { "goldenrod", SKColor.Parse("daa520") },
                { "goldenrod1", SKColor.Parse("ffc125") },
                { "goldenrod2", SKColor.Parse("eeb422") },
                { "goldenrod3", SKColor.Parse("cd9b1d") },
                { "goldenrod4", SKColor.Parse("8b6914") },
                { "gray", SKColor.Parse("bebebe") },
                { "gray1", SKColor.Parse("030303") },
                { "gray2", SKColor.Parse("050505") },
                { "gray3", SKColor.Parse("080808") },
                { "gray4", SKColor.Parse("0a0a0a") },
                { "gray5", SKColor.Parse("0d0d0d") },
                { "gray6", SKColor.Parse("0f0f0f") },
                { "gray7", SKColor.Parse("121212") },
                { "gray8", SKColor.Parse("141414") },
                { "gray9", SKColor.Parse("171717") },
                { "gray10", SKColor.Parse("1a1a1a") },
                { "gray11", SKColor.Parse("1c1c1c") },
                { "gray12", SKColor.Parse("1f1f1f") },
                { "gray13", SKColor.Parse("212121") },
                { "gray14", SKColor.Parse("242424") },
                { "gray15", SKColor.Parse("262626") },
                { "gray16", SKColor.Parse("292929") },
                { "gray17", SKColor.Parse("2b2b2b") },
                { "gray18", SKColor.Parse("2e2e2e") },
                { "gray19", SKColor.Parse("303030") },
                { "gray20", SKColor.Parse("333333") },
                { "gray21", SKColor.Parse("363636") },
                { "gray22", SKColor.Parse("383838") },
                { "gray23", SKColor.Parse("3b3b3b") },
                { "gray24", SKColor.Parse("3d3d3d") },
                { "gray25", SKColor.Parse("404040") },
                { "gray26", SKColor.Parse("424242") },
                { "gray27", SKColor.Parse("454545") },
                { "gray28", SKColor.Parse("474747") },
                { "gray29", SKColor.Parse("4a4a4a") },
                { "gray30", SKColor.Parse("4d4d4d") },
                { "gray31", SKColor.Parse("4f4f4f") },
                { "gray32", SKColor.Parse("525252") },
                { "gray33", SKColor.Parse("545454") },
                { "gray34", SKColor.Parse("575757") },
                { "gray35", SKColor.Parse("595959") },
                { "gray36", SKColor.Parse("5c5c5c") },
                { "gray37", SKColor.Parse("5e5e5e") },
                { "gray38", SKColor.Parse("616161") },
                { "gray39", SKColor.Parse("636363") },
                { "gray40", SKColor.Parse("666666") },
                { "gray41", SKColor.Parse("696969") },
                { "gray42", SKColor.Parse("6b6b6b") },
                { "gray43", SKColor.Parse("6e6e6e") },
                { "gray44", SKColor.Parse("707070") },
                { "gray45", SKColor.Parse("737373") },
                { "gray46", SKColor.Parse("757575") },
                { "gray47", SKColor.Parse("787878") },
                { "gray48", SKColor.Parse("7a7a7a") },
                { "gray49", SKColor.Parse("7d7d7d") },
                { "gray50", SKColor.Parse("7f7f7f") },
                { "gray51", SKColor.Parse("828282") },
                { "gray52", SKColor.Parse("858585") },
                { "gray53", SKColor.Parse("878787") },
                { "gray54", SKColor.Parse("8a8a8a") },
                { "gray55", SKColor.Parse("8c8c8c") },
                { "gray56", SKColor.Parse("8f8f8f") },
                { "gray57", SKColor.Parse("919191") },
                { "gray58", SKColor.Parse("949494") },
                { "gray59", SKColor.Parse("969696") },
                { "gray60", SKColor.Parse("999999") },
                { "gray61", SKColor.Parse("9c9c9c") },
                { "gray62", SKColor.Parse("9e9e9e") },
                { "gray63", SKColor.Parse("a1a1a1") },
                { "gray64", SKColor.Parse("a3a3a3") },
                { "gray65", SKColor.Parse("a6a6a6") },
                { "gray66", SKColor.Parse("a8a8a8") },
                { "gray67", SKColor.Parse("ababab") },
                { "gray68", SKColor.Parse("adadad") },
                { "gray69", SKColor.Parse("b0b0b0") },
                { "gray70", SKColor.Parse("b3b3b3") },
                { "gray71", SKColor.Parse("b5b5b5") },
                { "gray72", SKColor.Parse("b8b8b8") },
                { "gray73", SKColor.Parse("bababa") },
                { "gray74", SKColor.Parse("bdbdbd") },
                { "gray75", SKColor.Parse("bfbfbf") },
                { "gray76", SKColor.Parse("c2c2c2") },
                { "gray77", SKColor.Parse("c4c4c4") },
                { "gray78", SKColor.Parse("c7c7c7") },
                { "gray79", SKColor.Parse("c9c9c9") },
                { "gray80", SKColor.Parse("cccccc") },
                { "gray81", SKColor.Parse("cfcfcf") },
                { "gray82", SKColor.Parse("d1d1d1") },
                { "gray83", SKColor.Parse("d4d4d4") },
                { "gray84", SKColor.Parse("d6d6d6") },
                { "gray85", SKColor.Parse("d9d9d9") },
                { "gray86", SKColor.Parse("dbdbdb") },
                { "gray87", SKColor.Parse("dedede") },
                { "gray88", SKColor.Parse("e0e0e0") },
                { "gray89", SKColor.Parse("e3e3e3") },
                { "gray90", SKColor.Parse("e5e5e5") },
                { "gray91", SKColor.Parse("e8e8e8") },
                { "gray92", SKColor.Parse("ebebeb") },
                { "gray93", SKColor.Parse("ededed") },
                { "gray94", SKColor.Parse("f0f0f0") },
                { "gray95", SKColor.Parse("f2f2f2") },
                { "gray97", SKColor.Parse("f7f7f7") },
                { "gray98", SKColor.Parse("fafafa") },
                { "gray99", SKColor.Parse("fcfcfc") },
                { "green1", SKColor.Parse("00ff00") },
                { "green2", SKColor.Parse("00ee00") },
                { "green3", SKColor.Parse("00cd00") },
                { "green4", SKColor.Parse("008b00") },
                { "greenyellow", SKColor.Parse("adff2f") },
                { "honeydew1", SKColor.Parse("f0fff0") },
                { "honeydew2", SKColor.Parse("e0eee0") },
                { "honeydew3", SKColor.Parse("c1cdc1") },
                { "honeydew4", SKColor.Parse("838b83") },
                { "hotpink", SKColor.Parse("ff69b4") },
                { "hotpink1", SKColor.Parse("ff6eb4") },
                { "hotpink2", SKColor.Parse("ee6aa7") },
                { "hotpink3", SKColor.Parse("cd6090") },
                { "hotpink4", SKColor.Parse("8b3a62") },
                { "indianred", SKColor.Parse("cd5c5c") },
                { "indianred1", SKColor.Parse("ff6a6a") },
                { "indianred2", SKColor.Parse("ee6363") },
                { "indianred3", SKColor.Parse("cd5555") },
                { "indianred4", SKColor.Parse("8b3a3a") },
                { "ivory1", SKColor.Parse("fffff0") },
                { "ivory2", SKColor.Parse("eeeee0") },
                { "ivory3", SKColor.Parse("cdcdc1") },
                { "ivory4", SKColor.Parse("8b8b83") },
                { "khaki", SKColor.Parse("f0e68c") },
                { "khaki1", SKColor.Parse("fff68f") },
                { "khaki2", SKColor.Parse("eee685") },
                { "khaki3", SKColor.Parse("cdc673") },
                { "khaki4", SKColor.Parse("8b864e") },
                { "lavender", SKColor.Parse("e6e6fa") },
                { "lavenderblush1", SKColor.Parse("fff0f5") },
                { "lavenderblush2", SKColor.Parse("eee0e5") },
                { "lavenderblush3", SKColor.Parse("cdc1c5") },
                { "lavenderblush4", SKColor.Parse("8b8386") },
                { "lawngreen", SKColor.Parse("7cfc00") },
                { "lemonchiffon1", SKColor.Parse("fffacd") },
                { "lemonchiffon2", SKColor.Parse("eee9bf") },
                { "lemonchiffon3", SKColor.Parse("cdc9a5") },
                { "lemonchiffon4", SKColor.Parse("8b8970") },
                { "light", SKColor.Parse("eedd82") },
                { "lightblue", SKColor.Parse("add8e6") },
                { "lightblue1", SKColor.Parse("bfefff") },
                { "lightblue2", SKColor.Parse("b2dfee") },
                { "lightblue3", SKColor.Parse("9ac0cd") },
                { "lightblue4", SKColor.Parse("68838b") },
                { "lightcoral", SKColor.Parse("f08080") },
                { "lightcyan1", SKColor.Parse("e0ffff") },
                { "lightcyan2", SKColor.Parse("d1eeee") },
                { "lightcyan3", SKColor.Parse("b4cdcd") },
                { "lightcyan4", SKColor.Parse("7a8b8b") },
                { "lightgoldenrod1", SKColor.Parse("ffec8b") },
                { "lightgoldenrod2", SKColor.Parse("eedc82") },
                { "lightgoldenrod3", SKColor.Parse("cdbe70") },
                { "lightgoldenrod4", SKColor.Parse("8b814c") },
                { "lightgoldenrodyellow", SKColor.Parse("fafad2") },
                { "lightgray", SKColor.Parse("d3d3d3") },
                { "lightpink", SKColor.Parse("ffb6c1") },
                { "lightpink1", SKColor.Parse("ffaeb9") },
                { "lightpink2", SKColor.Parse("eea2ad") },
                { "lightpink3", SKColor.Parse("cd8c95") },
                { "lightpink4", SKColor.Parse("8b5f65") },
                { "lightsalmon1", SKColor.Parse("ffa07a") },
                { "lightsalmon2", SKColor.Parse("ee9572") },
                { "lightsalmon3", SKColor.Parse("cd8162") },
                { "lightsalmon4", SKColor.Parse("8b5742") },
                { "lightseagreen", SKColor.Parse("20b2aa") },
                { "lightskyblue", SKColor.Parse("87cefa") },
                { "lightskyblue1", SKColor.Parse("b0e2ff") },
                { "lightskyblue2", SKColor.Parse("a4d3ee") },
                { "lightskyblue3", SKColor.Parse("8db6cd") },
                { "lightskyblue4", SKColor.Parse("607b8b") },
                { "lightslateblue", SKColor.Parse("8470ff") },
                { "lightslategray", SKColor.Parse("778899") },
                { "lightsteelblue", SKColor.Parse("b0c4de") },
                { "lightsteelblue1", SKColor.Parse("cae1ff") },
                { "lightsteelblue2", SKColor.Parse("bcd2ee") },
                { "lightsteelblue3", SKColor.Parse("a2b5cd") },
                { "lightsteelblue4", SKColor.Parse("6e7b8b") },
                { "lightyellow1", SKColor.Parse("ffffe0") },
                { "lightyellow2", SKColor.Parse("eeeed1") },
                { "lightyellow3", SKColor.Parse("cdcdb4") },
                { "lightyellow4", SKColor.Parse("8b8b7a") },
                { "limegreen", SKColor.Parse("32cd32") },
                { "linen", SKColor.Parse("faf0e6") },
                { "magenta", SKColor.Parse("ff00ff") },
                { "magenta2", SKColor.Parse("ee00ee") },
                { "magenta3", SKColor.Parse("cd00cd") },
                { "magenta4", SKColor.Parse("8b008b") },
                { "maroon", SKColor.Parse("b03060") },
                { "maroon1", SKColor.Parse("ff34b3") },
                { "maroon2", SKColor.Parse("ee30a7") },
                { "maroon3", SKColor.Parse("cd2990") },
                { "maroon4", SKColor.Parse("8b1c62") },
                { "medium", SKColor.Parse("66cdaa") },
                { "mediumaquamarine", SKColor.Parse("66cdaa") },
                { "mediumblue", SKColor.Parse("0000cd") },
                { "mediumorchid", SKColor.Parse("ba55d3") },
                { "mediumorchid1", SKColor.Parse("e066ff") },
                { "mediumorchid2", SKColor.Parse("d15fee") },
                { "mediumorchid3", SKColor.Parse("b452cd") },
                { "mediumorchid4", SKColor.Parse("7a378b") },
                { "mediumpurple", SKColor.Parse("9370db") },
                { "mediumpurple1", SKColor.Parse("ab82ff") },
                { "mediumpurple2", SKColor.Parse("9f79ee") },
                { "mediumpurple3", SKColor.Parse("8968cd") },
                { "mediumpurple4", SKColor.Parse("5d478b") },
                { "mediumseagreen", SKColor.Parse("3cb371") },
                { "mediumslateblue", SKColor.Parse("7b68ee") },
                { "mediumspringgreen", SKColor.Parse("00fa9a") },
                { "mediumturquoise", SKColor.Parse("48d1cc") },
                { "mediumvioletred", SKColor.Parse("c71585") },
                { "midnightblue", SKColor.Parse("191970") },
                { "mintcream", SKColor.Parse("f5fffa") },
                { "mistyrose1", SKColor.Parse("ffe4e1") },
                { "mistyrose2", SKColor.Parse("eed5d2") },
                { "mistyrose3", SKColor.Parse("cdb7b5") },
                { "mistyrose4", SKColor.Parse("8b7d7b") },
                { "moccasin", SKColor.Parse("ffe4b5") },
                { "navajowhite1", SKColor.Parse("ffdead") },
                { "navajowhite2", SKColor.Parse("eecfa1") },
                { "navajowhite3", SKColor.Parse("cdb38b") },
                { "navajowhite4", SKColor.Parse("8b795e") },
                { "navyblue", SKColor.Parse("000080") },
                { "oldlace", SKColor.Parse("fdf5e6") },
                { "olivedrab", SKColor.Parse("6b8e23") },
                { "olivedrab1", SKColor.Parse("c0ff3e") },
                { "olivedrab2", SKColor.Parse("b3ee3a") },
                { "olivedrab4", SKColor.Parse("698b22") },
                { "orange1", SKColor.Parse("ffa500") },
                { "orange2", SKColor.Parse("ee9a00") },
                { "orange3", SKColor.Parse("cd8500") },
                { "orange4", SKColor.Parse("8b5a00") },
                { "orangered1", SKColor.Parse("ff4500") },
                { "orangered2", SKColor.Parse("ee4000") },
                { "orangered3", SKColor.Parse("cd3700") },
                { "orangered4", SKColor.Parse("8b2500") },
                { "orchid", SKColor.Parse("da70d6") },
                { "orchid1", SKColor.Parse("ff83fa") },
                { "orchid2", SKColor.Parse("ee7ae9") },
                { "orchid3", SKColor.Parse("cd69c9") },
                { "orchid4", SKColor.Parse("8b4789") },
                { "pale", SKColor.Parse("db7093") },
                { "palegoldenrod", SKColor.Parse("eee8aa") },
                { "palegreen", SKColor.Parse("98fb98") },
                { "palegreen1", SKColor.Parse("9aff9a") },
                { "palegreen2", SKColor.Parse("90ee90") },
                { "palegreen3", SKColor.Parse("7ccd7c") },
                { "palegreen4", SKColor.Parse("548b54") },
                { "paleturquoise", SKColor.Parse("afeeee") },
                { "paleturquoise1", SKColor.Parse("bbffff") },
                { "paleturquoise2", SKColor.Parse("aeeeee") },
                { "paleturquoise3", SKColor.Parse("96cdcd") },
                { "paleturquoise4", SKColor.Parse("668b8b") },
                { "palevioletred", SKColor.Parse("db7093") },
                { "palevioletred1", SKColor.Parse("ff82ab") },
                { "palevioletred2", SKColor.Parse("ee799f") },
                { "palevioletred3", SKColor.Parse("cd6889") },
                { "palevioletred4", SKColor.Parse("8b475d") },
                { "papayawhip", SKColor.Parse("ffefd5") },
                { "peachpuff1", SKColor.Parse("ffdab9") },
                { "peachpuff2", SKColor.Parse("eecbad") },
                { "peachpuff3", SKColor.Parse("cdaf95") },
                { "peachpuff4", SKColor.Parse("8b7765") },
                { "pink", SKColor.Parse("ffc0cb") },
                { "pink1", SKColor.Parse("ffb5c5") },
                { "pink2", SKColor.Parse("eea9b8") },
                { "pink3", SKColor.Parse("cd919e") },
                { "pink4", SKColor.Parse("8b636c") },
                { "plum", SKColor.Parse("dda0dd") },
                { "plum1", SKColor.Parse("ffbbff") },
                { "plum2", SKColor.Parse("eeaeee") },
                { "plum3", SKColor.Parse("cd96cd") },
                { "plum4", SKColor.Parse("8b668b") },
                { "powderblue", SKColor.Parse("b0e0e6") },
                { "purple", SKColor.Parse("a020f0") },
                { "purple1", SKColor.Parse("9b30ff") },
                { "purple2", SKColor.Parse("912cee") },
                { "purple3", SKColor.Parse("7d26cd") },
                { "purple4", SKColor.Parse("551a8b") },
                { "red1", SKColor.Parse("ff0000") },
                { "red2", SKColor.Parse("ee0000") },
                { "red3", SKColor.Parse("cd0000") },
                { "red4", SKColor.Parse("8b0000") },
                { "rosybrown", SKColor.Parse("bc8f8f") },
                { "rosybrown1", SKColor.Parse("ffc1c1") },
                { "rosybrown2", SKColor.Parse("eeb4b4") },
                { "rosybrown3", SKColor.Parse("cd9b9b") },
                { "rosybrown4", SKColor.Parse("8b6969") },
                { "royalblue", SKColor.Parse("4169e1") },
                { "royalblue1", SKColor.Parse("4876ff") },
                { "royalblue2", SKColor.Parse("436eee") },
                { "royalblue3", SKColor.Parse("3a5fcd") },
                { "royalblue4", SKColor.Parse("27408b") },
                { "saddlebrown", SKColor.Parse("8b4513") },
                { "salmon", SKColor.Parse("fa8072") },
                { "salmon1", SKColor.Parse("ff8c69") },
                { "salmon2", SKColor.Parse("ee8262") },
                { "salmon3", SKColor.Parse("cd7054") },
                { "salmon4", SKColor.Parse("8b4c39") },
                { "sandybrown", SKColor.Parse("f4a460") },
                { "seagreen1", SKColor.Parse("54ff9f") },
                { "seagreen2", SKColor.Parse("4eee94") },
                { "seagreen3", SKColor.Parse("43cd80") },
                { "seagreen4", SKColor.Parse("2e8b57") },
                { "seashell1", SKColor.Parse("fff5ee") },
                { "seashell2", SKColor.Parse("eee5de") },
                { "seashell3", SKColor.Parse("cdc5bf") },
                { "seashell4", SKColor.Parse("8b8682") },
                { "sienna", SKColor.Parse("a0522d") },
                { "sienna1", SKColor.Parse("ff8247") },
                { "sienna2", SKColor.Parse("ee7942") },
                { "sienna3", SKColor.Parse("cd6839") },
                { "sienna4", SKColor.Parse("8b4726") },
                { "skyblue", SKColor.Parse("87ceeb") },
                { "skyblue1", SKColor.Parse("87ceff") },
                { "skyblue2", SKColor.Parse("7ec0ee") },
                { "skyblue3", SKColor.Parse("6ca6cd") },
                { "skyblue4", SKColor.Parse("4a708b") },
                { "slateblue", SKColor.Parse("6a5acd") },
                { "slateblue1", SKColor.Parse("836fff") },
                { "slateblue2", SKColor.Parse("7a67ee") },
                { "slateblue3", SKColor.Parse("6959cd") },
                { "slateblue4", SKColor.Parse("473c8b") },
                { "slategray", SKColor.Parse("708090") },
                { "slategray1", SKColor.Parse("c6e2ff") },
                { "slategray2", SKColor.Parse("b9d3ee") },
                { "slategray3", SKColor.Parse("9fb6cd") },
                { "slategray4", SKColor.Parse("6c7b8b") },
                { "snow1", SKColor.Parse("fffafa") },
                { "snow2", SKColor.Parse("eee9e9") },
                { "snow3", SKColor.Parse("cdc9c9") },
                { "snow4", SKColor.Parse("8b8989") },
                { "springgreen1", SKColor.Parse("00ff7f") },
                { "springgreen2", SKColor.Parse("00ee76") },
                { "springgreen3", SKColor.Parse("00cd66") },
                { "springgreen4", SKColor.Parse("008b45") },
                { "steelblue", SKColor.Parse("4682b4") },
                { "steelblue1", SKColor.Parse("63b8ff") },
                { "steelblue2", SKColor.Parse("5cacee") },
                { "steelblue3", SKColor.Parse("4f94cd") },
                { "steelblue4", SKColor.Parse("36648b") },
                { "tan", SKColor.Parse("d2b48c") },
                { "tan1", SKColor.Parse("ffa54f") },
                { "tan2", SKColor.Parse("ee9a49") },
                { "tan3", SKColor.Parse("cd853f") },
                { "tan4", SKColor.Parse("8b5a2b") },
                { "thistle", SKColor.Parse("d8bfd8") },
                { "thistle1", SKColor.Parse("ffe1ff") },
                { "thistle2", SKColor.Parse("eed2ee") },
                { "thistle3", SKColor.Parse("cdb5cd") },
                { "thistle4", SKColor.Parse("8b7b8b") },
                { "tomato1", SKColor.Parse("ff6347") },
                { "tomato2", SKColor.Parse("ee5c42") },
                { "tomato3", SKColor.Parse("cd4f39") },
                { "tomato4", SKColor.Parse("8b3626") },
                { "turquoise", SKColor.Parse("40e0d0") },
                { "turquoise1", SKColor.Parse("00f5ff") },
                { "turquoise2", SKColor.Parse("00e5ee") },
                { "turquoise3", SKColor.Parse("00c5cd") },
                { "turquoise4", SKColor.Parse("00868b") },
                { "violet", SKColor.Parse("ee82ee") },
                { "violetred", SKColor.Parse("d02090") },
                { "violetred1", SKColor.Parse("ff3e96") },
                { "violetred2", SKColor.Parse("ee3a8c") },
                { "violetred3", SKColor.Parse("cd3278") },
                { "violetred4", SKColor.Parse("8b2252") },
                { "wheat", SKColor.Parse("f5deb3") },
                { "wheat1", SKColor.Parse("ffe7ba") },
                { "wheat2", SKColor.Parse("eed8ae") },
                { "wheat3", SKColor.Parse("cdba96") },
                { "wheat4", SKColor.Parse("8b7e66") },
                { "white", SKColor.Parse("ffffff") },
                { "whitesmoke", SKColor.Parse("f5f5f5") },
                { "yellow1", SKColor.Parse("ffff00") },
                { "yellow2", SKColor.Parse("eeee00") },
                { "yellow3", SKColor.Parse("cdcd00") },
                { "yellow4", SKColor.Parse("8b8b00") },
                { "yellowgreen", SKColor.Parse("9acd32") }
            };

            foreach (var field in typeof(SKColors).GetFields(BindingFlags.Static | BindingFlags.Public))
            {
                if (!_library.ContainsKey(field.Name))
                {
                    _library[field.Name] = (SKColor)field.GetValue(null);
                }
            }
        }
    }
}
