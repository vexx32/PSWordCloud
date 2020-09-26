using System;
using System.Collections;
using System.Collections.Generic;
using System.Management.Automation;
using System.Text;
using System.Xml;
using SkiaSharp;

namespace PSWordCloud
{
    internal static class Extensions
    {
        internal static SKPoint Multiply(this SKPoint point, float factor)
            => new SKPoint(point.X * factor, point.Y * factor);

        internal static float ToRadians(this float degrees)
            => (float)(degrees * Math.PI / 180);

        /// <summary>
        /// Utility method which is just a convenient shortcut to <see cref="LanguagePrimitives.ConvertTo{T}(object)"/>.
        /// </summary>
        /// <param name="item">The object to convert.</param>
        /// <typeparam name="TSource">The original object type.</typeparam>
        /// <typeparam name="TResult">The resulting destination type.</typeparam>
        /// <returns>The converted value.</returns>
        public static TResult ConvertTo<TResult>(this object item)
            => LanguagePrimitives.ConvertTo<TResult>(item);

        /// <summary>
        /// Perform an in-place-modification operation on every element in the array.
        /// </summary>
        /// <param name="items">The array to operate on.</param>
        /// <returns>The transformed array.</returns>
        public static T[] TransformElements<T>(this T[] items, Func<T, T> operation)
        {
            for (var index = 0; index < items.Length; index++)
            {
                items[index] = operation.Invoke(items[index]);
            }

            return items;
        }

        public static SKColor AsMonochrome(this SKColor color)
        {
            color.ToHsv(out _, out _, out float brightness);
            byte level = (byte)Math.Floor(255 * brightness / 100f);

            return new SKColor(level, level, level);
        }

        /// <summary>
        /// Determines whether a given color is considered sufficiently visually distinct from a backdrop color.
        /// </summary>
        /// <param name="target">The target color.</param>
        /// <param name="backdrop">A reference color to compare against.</param>
        internal static bool IsDistinctFrom(this SKColor target, SKColor backdrop)
        {
            if (target.IsTransparent())
            {
                return false;
            }

            backdrop.ToHsv(out float refHue, out float refSaturation, out float refBrightness);
            target.ToHsv(out float hue, out float saturation, out float brightness);

            float brightnessDistance = Math.Abs(refBrightness - brightness);
            if (brightnessDistance > 30)
            {
                return true;
            }

            float hueDistance = Math.Abs(refHue - hue);
            if (hueDistance > 24 && brightnessDistance > 20)
            {
                return true;
            }

            float saturationDistance = Math.Abs(refSaturation - saturation);
            return saturationDistance > 24 && brightnessDistance > 18;
        }

        private static bool IsTransparent(this SKColor color) => color.Alpha == 0;

        /// <summary>
        /// Gets the smallest square that would completely contain the given <paramref name="rectangle"/>, with the rectangle positioned
        /// at its centre.
        /// </summary>
        /// <param name="rectangle">The rectangle to find the containing square for.</param>
        internal static SKRect GetEnclosingSquare(this SKRect rectangle)
            => rectangle switch
            {
                SKRect wide when wide.Width > wide.Height => SKRect.Inflate(wide, x: 0, y: (wide.Width - wide.Height) / 2),
                SKRect tall when tall.Height > tall.Width => SKRect.Inflate(tall, x: (tall.Height - tall.Width) / 2, y: 0),
                _ => SKRect.Inflate(rectangle, x: 0, y: 0)
            };

        /// <summary>
        /// Returns a random <see cref="float"/> value between the specified minimum and maximum.
        /// </summary>
        /// <param name="random">The <see cref="Random"/> instance to use for the generation.</param>
        /// <param name="min">The minimum float value.</param>
        /// <param name="max">The maximum float value.</param>
        internal static float NextFloat(this Random random, float min, float max)
        {
            if (min > max)
            {
                return max;
            }

            var range = max - min;
            return (float)random.NextDouble() * range + min;
        }

        /// <summary>
        /// Performs an in-place random shuffle on an array by swapping elements.
        /// This algorithm is pretty commonly used, but effective and fast enough for our purposes.
        /// </summary>
        /// <param name="rng">Random number generator.</param>
        /// <param name="array">The array to shuffle.</param>
        /// <typeparam name="T">The element type of the array.</typeparam>
        internal static IList<T> Shuffle<T>(this Random rng, IList<T> array)
        {
            int n = array.Count;
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
        /// <param name="rectangle">The rectangle to test position against the edges of the region.</param>
        /// <param name="region">The region to test against.</param>
        /// <returns>Returns false if the rectangle is entirely within the region, and false otherwise.</returns>
        internal static bool FallsOutside(this SKRect rectangle, SKRegion region)
        {
            var bounds = region.Bounds;
            return rectangle.Top < bounds.Top
                || rectangle.Bottom > bounds.Bottom
                || rectangle.Left < bounds.Left
                || rectangle.Right > bounds.Right;
        }

        /// <summary>
        /// Checks if the given point is outside the given <see cref="SKRegion"/>.
        /// </summary>
        /// <param name="point">The point in question.</param>
        /// <param name="region">The region to check against.</param>
        /// <returns>Returns true if the point is within the given region.</returns>
        internal static bool IsOutside(this SKPoint point, SKRegion region) => !region.Contains(point);

        /// <summary>
        /// Translates the <see cref="SKPath"/> so that its midpoint is the given position.
        /// </summary>
        /// <param name="path">The path to translate.</param>
        /// <param name="point">The new point to centre the path on.</param>
        internal static void CentreOnPoint(this SKPath path, SKPoint point)
        {
            var pathMidpoint = new SKPoint(path.TightBounds.MidX, path.TightBounds.MidY);
            path.Offset(point - pathMidpoint);
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

        internal static void SetFill(this SKPaint brush, SKColor fill)
        {
            brush.IsStroke = false;
            brush.Style = SKPaintStyle.Fill;
            brush.Color = fill;
        }

        internal static void SetStroke(this SKPaint brush, SKColor stroke, float width)
        {
            brush.IsStroke = true;
            brush.StrokeWidth = width;
            brush.Style = SKPaintStyle.Stroke;
            brush.Color = stroke;
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
        /// Rotates the given path by the declared angle.
        /// </summary>
        /// <param name="path">The path to rotate.</param>
        /// <param name="degrees">The angle in degrees to rotate the path.</param>
        internal static void Rotate(this SKPath path, float degrees)
        {
            SKRect pathBounds = path.TightBounds;
            SKMatrix rotation = SKMatrix.CreateRotationDegrees(degrees, pathBounds.MidX, pathBounds.MidY);
            path.Transform(rotation);
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

            return region.Intersects(SKRectI.Round(rect));
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

            return region.Intersects(path);
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

        internal static object? GetValue(this IEnumerable collection, string key)
            => collection switch
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
}
