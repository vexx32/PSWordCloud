using System;
using SkiaSharp;

namespace PSWordCloud
{
    static class Utils
    {
        public static SKRectI ToSKRectI(this SKRect rectangle)
        {
            return new SKRectI((int)rectangle.Left, (int)rectangle.Top, (int)rectangle.Right, (int)rectangle.Bottom);
        }
    }
}
