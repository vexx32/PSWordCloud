using System;
using System.Linq;
using Xunit;
using PSWordCloud;
using SkiaSharp;

namespace PSWordCloud.Tests
{
    public class WCUtilsTests
    {
        [Fact]
        public void Test_ToRadians()
        {
            //Assert.Equal(KnownResult, TestResult);

            Assert.Equal(0, ((float)0).ToRadians());
        }

        [Fact]
        public void Test_Multiply()
        {
            var point = new SKPoint(1, 2);

            Assert.Equal(point.Multiply(5), new SKPoint(5, 10));
            Assert.Equal(point.Multiply(1.5f), new SKPoint(1.5f, 3));
        }

        [Fact]
        public void Test_Shuffle()
        {
            var collection = new[] { 1, 2, 3, 4, 5 };
            var original = new[] { 1, 2, 3, 4, 5 };
            var rng = new Random();
            rng.Shuffle(collection);

            Assert.False(Enumerable.SequenceEqual(original, collection));
        }

        [Fact]
        public void Test_FallsOutside()
        {
            var rect = new SKRect(0, 0, 10, 10);
            var region = new SKRegion();

            region.SetRect(new SKRectI(-5, -5, 5, 5));

            Assert.True(rect.FallsOutside(region));

            rect = new SKRect(-5, -5, 5, 5);
            Assert.False(rect.FallsOutside(region));
        }

        [Fact]
        public void Test_NextWord()
        {
            var brush = new SKPaint();
            brush.NextWord(2.4f, 1.8f, SKColors.Beige);

            Assert.Equal(brush.TextSize, 2.4f);
            Assert.False(brush.IsStroke);
            Assert.Equal(brush.Style, SKPaintStyle.StrokeAndFill);
            Assert.Equal(brush.StrokeWidth, 1.8f * 1.8f * NewWordCloudCommand.STROKE_BASE_SCALE);
            Assert.False(brush.IsVerticalText);
            Assert.Equal(brush.Color, SKColors.Beige);
        }

        [Fact]
        public void Test_SortValue()
        {
            var color = SKColors.Blue;
            var expectedValue = 100 * (1 - 0.5) / (1 - 100);

            Assert.Equal(color.SortValue(1), expectedValue);
        }
    }
}
