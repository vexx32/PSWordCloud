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
    }
}
