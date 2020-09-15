using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Management.Automation;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
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
        /// <summary>
        /// Returns a font scale value based on the size of the letter X in a given typeface when the text size is 1 unit.
        /// </summary>
        /// <param name="typeface">The typeface to measure the scale from.</param>
        /// <returns>A float value representing the area occupied by the letter X in a typeface.</returns>
        internal static float GetAverageCharArea(SKTypeface typeface)
        {
            const string alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz";
            using SKPaint brush = GetBrush(wordSize: 1, strokeWidth: 0, typeface: typeface);

            using SKPath textPath = brush.GetTextPath(alphabet, x: 0, y: 0);

            SKRect rect = textPath.TightBounds;
            return rect.Width * rect.Height / alphabet.Length;
        }

        /// <summary>
        /// Returns a shuffled set of possible angles determined by the <paramref name="permittedRotations"/>.
        /// </summary>
        internal static IReadOnlyList<float> GetDrawAngles(WordOrientations permittedRotations, LockingRandom random)
            => (IReadOnlyList<float>)(permittedRotations switch
            {
                WordOrientations.Vertical => random.Shuffle(new float[] { 0, 90 }),
                WordOrientations.FlippedVertical => random.Shuffle(new float[] { 0, -90 }),
                WordOrientations.EitherVertical => random.Shuffle(new float[] { 0, 90, -90 }),
                WordOrientations.UprightDiagonals => random.Shuffle(new float[] { 0, -90, -45, 45, 90 }),
                WordOrientations.InvertedDiagonals => random.Shuffle(new float[] { 90, 135, -135, -90, 180 }),
                WordOrientations.AllDiagonals => random.Shuffle(new float[] { 45, 90, 135, 180, -135, -90, -45, 0 }),
                WordOrientations.AllUpright => random.GetRandomFloats(-90, 91),
                WordOrientations.AllInverted => random.GetRandomFloats(90, 271),
                WordOrientations.All => random.GetRandomFloats(0, 361),
                _ => new float[] { 0 },
            });

        /// <summary>
        /// Prepares the brush to draw the next word.
        /// This overload assumes the text to be drawn will be black.
        /// </summary>
        /// <param name="wordSize"></param>
        /// <param name="strokeWidth"></param>
        /// <param name="typeface"></param>
        internal static SKPaint GetBrush(float wordSize, float strokeWidth, SKTypeface typeface)
            => GetBrush(wordSize, strokeWidth, SKColors.Black, typeface);

        /// <summary>
        /// Prepares the brush to draw the next word.
        /// </summary>
        /// <param name="brush">The brush to prepare.</param>
        /// <param name="wordSize">The size of the word we'll be drawing.</param>
        /// <param name="strokeWidth">Width of the stroke we'll be drawing.</param>
        /// <param name="color">Color of the word we'll be drawing.</param>
        /// <param name="typeface">The typeface to draw words with.</param>
        internal static SKPaint GetBrush(
            float wordSize,
            float strokeWidth,
            SKColor color,
            SKTypeface typeface)
            => new SKPaint
            {
                Typeface = typeface,
                TextSize = wordSize,
                Style = SKPaintStyle.StrokeAndFill,
                Color = color,
                StrokeWidth = wordSize * strokeWidth * Constants.StrokeBaseScale,
                IsStroke = false,
                IsAutohinted = true,
                IsAntialias = true
            };

        /// <summary>
        /// Gets the appropriate word bubble path for the requested shape, sized to fit the word bounds.
        /// </summary>
        /// <param name="shape">The shape of the bubble.</param>
        /// <param name="wordBounds">The bounds of the word to surround.</param>
        /// <returns>The <see cref="SKPath"/> representing the word bubble.</returns>
        internal static SKPath GetWordBubblePath(WordBubbleShape shape, SKRect wordBounds)
            => shape switch
            {
                WordBubbleShape.Rectangle => GetRectanglePath(wordBounds),
                WordBubbleShape.Square => GetSquarePath(wordBounds),
                WordBubbleShape.Circle => GetCirclePath(wordBounds),
                WordBubbleShape.Oval => GetOvalPath(wordBounds),
                _ => throw new ArgumentOutOfRangeException(nameof(shape))
            };

        private static SKPath GetRectanglePath(SKRect rectangle)
        {
            var path = new SKPath();
            float cornerRadius = rectangle.Height / 16;
            path.AddRoundRect(new SKRoundRect(rectangle, cornerRadius, cornerRadius));

            return path;
        }

        private static SKPath GetSquarePath(SKRect rectangle)
        {
            var path = new SKPath();
            float cornerRadius = Math.Max(rectangle.Width, rectangle.Height) / 16;
            path.AddRoundRect(new SKRoundRect(rectangle.GetEnclosingSquare(), cornerRadius, cornerRadius));

            return path;
        }

        private static SKPath GetCirclePath(SKRect rectangle)
        {
            var path = new SKPath();
            float bubbleRadius = Math.Max(rectangle.Width, rectangle.Height) / 2;
            path.AddCircle(rectangle.MidX, rectangle.MidY, bubbleRadius);

            return path;
        }

        private static SKPath GetOvalPath(SKRect rectangle)
        {
            var path = new SKPath();
            path.AddOval(rectangle);

            return path;
        }

        internal static bool AngleIsMostlyVertical(float degrees)
        {
            float remainder = Math.Abs(degrees % 180);
            return 135 > remainder && remainder > 45;
        }

        private static bool WordWillFit(SKRect wordBounds, SKRegion occupiedSpace)
            => !occupiedSpace.IntersectsRect(wordBounds);

        private static bool WordBubbleWillFit(
            WordBubbleShape shape,
            SKRect wordBounds,
            SKRegion occupiedSpace,
            out SKPath bubblePath)
        {
            bubblePath = GetWordBubblePath(shape, wordBounds);
            return !occupiedSpace.IntersectsPath(bubblePath);
        }

        /// <summary>
        /// Checks whether the given word bounds rectangle and the bubble surrounding it will fit in the desired
        /// location without bleeding over the <paramref name="clipRegion"/> or intersecting already-drawn words
        /// or their bubbles (which are recorded in the <paramref name="occupiedSpace"/> region).
        /// </summary>
        /// <param name="wordBounds">The rectangular bounds of the word to attempt to fit.</param>
        /// <param name="bubbleShape">The shape of the word bubble we'll need to draw.</param>
        /// <param name="clipRegion">The region that defines the allowable draw area.</param>
        /// <param name="occupiedSpace">The region that defines the space in the image that's already occupied.</param>
        /// <returns>Returns true if the word and its surrounding bubble have sufficient space to be drawn.</returns>
        internal static bool WordWillFit(
            SKRect wordBounds,
            WordBubbleShape bubbleShape,
            SKRegion clipRegion,
            SKRegion occupiedSpace,
            out SKPath? bubblePath)
        {
            bubblePath = null;
            if (wordBounds.FallsOutside(clipRegion))
            {
                return false;
            }

            if (bubbleShape == WordBubbleShape.None)
            {
                return WordWillFit(wordBounds, occupiedSpace);
            }

            return WordBubbleWillFit(bubbleShape, wordBounds, occupiedSpace, out bubblePath);
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
            => ColorLibrary[colorName];

        internal static SKFontManager FontManager = SKFontManager.Default;

        internal static IOrderedEnumerable<string> FontList = WCUtils.FontManager.FontFamilies
            .OrderBy(name => name, StringComparer.OrdinalIgnoreCase);

        internal static ReadOnlyDictionary<string, (string Tooltip, SKSizeI Size)> StandardImageSizes =
            new ReadOnlyDictionary<string, (string, SKSizeI)>(new Dictionary<string, (string, SKSizeI)>() {
                {"480x800",     ("Mobile Screen Size (small)",  new SKSizeI( 480,  800))},
                {"640x1146",    ("Mobile Screen Size (medium)", new SKSizeI( 640, 1146))},
                {"720p",        ("Standard HD 1280x720",        new SKSizeI(1280,  720))},
                {"1080p",       ("Full HD 1920x1080",           new SKSizeI(1920, 1080))},
                {"4K",          ("Ultra HD 3840x2160",          new SKSizeI(3840, 2160))},
                {"A4",          ("816x1056",                    new SKSizeI( 816, 1056))},
                {"Poster11x17", ("1056x1632",                   new SKSizeI(1056, 1632))},
                {"Poster18x24", ("1728x2304",                   new SKSizeI(1728, 2304))},
                {"Poster24x36", ("2304x3456",                   new SKSizeI(2304, 3456))},
            });

        internal static readonly char[] SplitChars= new[] {
            ' ','\n','\t','\r','.',',',';','\\','/','|',
            ':','"','?','!','{','}','[',']',':','(',')',
            '<','>','“','”','*','#','%','^','&','+','=' };

        internal static string[] StopWords= new[] {
            "a","about","above","after","again","against","all","am","an","and","any","are","aren't","as","at","be",
            "because","been","before","being","below","between","both","but","by","can't","cannot","could","couldn't",
            "did","didn't","do","does","doesn't","doing","don't","down","during","each","few","for","from","further",
            "had","hadn't","has","hasn't","have","haven't","having","he","he'd","he'll","he's","her","here","here's",
            "hers","herself","him","himself","his","how","how's","i","i'd","i'll","i'm","i've","if","in","into","is",
            "isn't","it","it's","its","itself","let's","me","more","most","mustn't","my","myself","no","nor","not","of",
            "off","on","once","only","or","other","ought","our","ours","ourselves","out","over","own","same","shan't",
            "she","she'd","she'll","she's","should","shouldn't","so","some","such","than","that","that's","the","their",
            "theirs","them","themselves","then","there","there's","these","they","they'd","they'll","they're","they've",
            "this","those","through","to","too","under","until","up","very","was","wasn't","we","we'd","we'll","we're",
            "we've","were","weren't","what","what's","when","when's","where","where's","which","while","who","who's",
            "whom","why","why's","with","won't","would","wouldn't","you","you'd","you'll","you're","you've","your",
            "yours","yourself","yourselves" };

        internal static bool IsStopWord(string word) => StopWords.Contains(word, StringComparer.OrdinalIgnoreCase);

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
                { "AliceBlue", SKColor.Parse("f0f8ff") },
                { "AntiqueWhite", SKColor.Parse("faebd7") },
                { "AntiqueWhite1", SKColor.Parse("ffefdb") },
                { "AntiqueWhite2", SKColor.Parse("eedfcc") },
                { "AntiqueWhite3", SKColor.Parse("cdc0b0") },
                { "AntiqueWhite4", SKColor.Parse("8b8378") },
                { "Aquamarine1", SKColor.Parse("7fffd4") },
                { "Aquamarine2", SKColor.Parse("76eec6") },
                { "Aquamarine4", SKColor.Parse("458b74") },
                { "Azure1", SKColor.Parse("f0ffff") },
                { "Azure2", SKColor.Parse("e0eeee") },
                { "Azure3", SKColor.Parse("c1cdcd") },
                { "Azure4", SKColor.Parse("838b8b") },
                { "Beige", SKColor.Parse("f5f5dc") },
                { "Bisque1", SKColor.Parse("ffe4c4") },
                { "Bisque2", SKColor.Parse("eed5b7") },
                { "Bisque3", SKColor.Parse("cdb79e") },
                { "Bisque4", SKColor.Parse("8b7d6b") },
                { "Black", SKColor.Parse("000000") },
                { "BlanchedAlmond", SKColor.Parse("ffebcd") },
                { "Blue1", SKColor.Parse("0000ff") },
                { "Blue2", SKColor.Parse("0000ee") },
                { "Blue4", SKColor.Parse("00008b") },
                { "BlueViolet", SKColor.Parse("8a2be2") },
                { "Brown", SKColor.Parse("a52a2a") },
                { "Brown1", SKColor.Parse("ff4040") },
                { "Brown2", SKColor.Parse("ee3b3b") },
                { "Brown3", SKColor.Parse("cd3333") },
                { "Brown4", SKColor.Parse("8b2323") },
                { "Burlywood", SKColor.Parse("deb887") },
                { "Burlywood1", SKColor.Parse("ffd39b") },
                { "Burlywood2", SKColor.Parse("eec591") },
                { "Burlywood3", SKColor.Parse("cdaa7d") },
                { "Burlywood4", SKColor.Parse("8b7355") },
                { "CadetBlue", SKColor.Parse("5f9ea0") },
                { "CadetBlue1", SKColor.Parse("98f5ff") },
                { "CadetBlue2", SKColor.Parse("8ee5ee") },
                { "CadetBlue3", SKColor.Parse("7ac5cd") },
                { "CadetBlue4", SKColor.Parse("53868b") },
                { "Chartreuse1", SKColor.Parse("7fff00") },
                { "Chartreuse2", SKColor.Parse("76ee00") },
                { "Chartreuse3", SKColor.Parse("66cd00") },
                { "Chartreuse4", SKColor.Parse("458b00") },
                { "Chocolate", SKColor.Parse("d2691e") },
                { "Chocolate1", SKColor.Parse("ff7f24") },
                { "Chocolate2", SKColor.Parse("ee7621") },
                { "Chocolate3", SKColor.Parse("cd661d") },
                { "Coral", SKColor.Parse("ff7f50") },
                { "Coral1", SKColor.Parse("ff7256") },
                { "Coral2", SKColor.Parse("ee6a50") },
                { "Coral3", SKColor.Parse("cd5b45") },
                { "Coral4", SKColor.Parse("8b3e2f") },
                { "CornflowerBlue", SKColor.Parse("6495ed") },
                { "Cornsilk1", SKColor.Parse("fff8dc") },
                { "Cornsilk2", SKColor.Parse("eee8cd") },
                { "Cornsilk3", SKColor.Parse("cdc8b1") },
                { "Cornsilk4", SKColor.Parse("8b8878") },
                { "Cyan1", SKColor.Parse("00ffff") },
                { "Cyan2", SKColor.Parse("00eeee") },
                { "Cyan3", SKColor.Parse("00cdcd") },
                { "Cyan4", SKColor.Parse("008b8b") },
                { "DarkGoldenrod", SKColor.Parse("b8860b") },
                { "DarkGoldenrod1", SKColor.Parse("ffb90f") },
                { "DarkGoldenrod2", SKColor.Parse("eead0e") },
                { "DarkGoldenrod3", SKColor.Parse("cd950c") },
                { "DarkGoldenrod4", SKColor.Parse("8b6508") },
                { "DarkGreen", SKColor.Parse("006400") },
                { "DarkKhaki", SKColor.Parse("bdb76b") },
                { "DarkOliveGreen", SKColor.Parse("556b2f") },
                { "DarkOliveGreen1", SKColor.Parse("caff70") },
                { "DarkOliveGreen2", SKColor.Parse("bcee68") },
                { "DarkOliveGreen3", SKColor.Parse("a2cd5a") },
                { "DarkOliveGreen4", SKColor.Parse("6e8b3d") },
                { "DarkOrange", SKColor.Parse("ff8c00") },
                { "DarkOrange1", SKColor.Parse("ff7f00") },
                { "DarkOrange2", SKColor.Parse("ee7600") },
                { "DarkOrange3", SKColor.Parse("cd6600") },
                { "DarkOrange4", SKColor.Parse("8b4500") },
                { "DarkOrchid", SKColor.Parse("9932cc") },
                { "DarkOrchid1", SKColor.Parse("bf3eff") },
                { "DarkOrchid2", SKColor.Parse("b23aee") },
                { "DarkOrchid3", SKColor.Parse("9a32cd") },
                { "DarkOrchid4", SKColor.Parse("68228b") },
                { "DarkSalmon", SKColor.Parse("e9967a") },
                { "DarkSeaGreen", SKColor.Parse("8fbc8f") },
                { "DarkSeaGreen1", SKColor.Parse("c1ffc1") },
                { "DarkSeaGreen2", SKColor.Parse("b4eeb4") },
                { "DarkSeaGreen3", SKColor.Parse("9bcd9b") },
                { "DarkSeaGreen4", SKColor.Parse("698b69") },
                { "DarkSlateBlue", SKColor.Parse("483d8b") },
                { "DarkSlateGray", SKColor.Parse("2f4f4f") },
                { "DarkSlateGray1", SKColor.Parse("97ffff") },
                { "DarkSlateGray2", SKColor.Parse("8deeee") },
                { "DarkSlateGray3", SKColor.Parse("79cdcd") },
                { "DarkSlateGray4", SKColor.Parse("528b8b") },
                { "DarkTurquoise", SKColor.Parse("00ced1") },
                { "DarkViolet", SKColor.Parse("9400d3") },
                { "DeepPink1", SKColor.Parse("ff1493") },
                { "DeepPink2", SKColor.Parse("ee1289") },
                { "DeepPink3", SKColor.Parse("cd1076") },
                { "DeepPink4", SKColor.Parse("8b0a50") },
                { "DeepSkyBlue1", SKColor.Parse("00bfff") },
                { "DeepSkyBlue2", SKColor.Parse("00b2ee") },
                { "DeepSkyBlue3", SKColor.Parse("009acd") },
                { "DeepSkyBlue4", SKColor.Parse("00688b") },
                { "DimGray", SKColor.Parse("696969") },
                { "DodgerBlue1", SKColor.Parse("1e90ff") },
                { "DodgerBlue2", SKColor.Parse("1c86ee") },
                { "DodgerBlue3", SKColor.Parse("1874cd") },
                { "DodgerBlue4", SKColor.Parse("104e8b") },
                { "Firebrick", SKColor.Parse("b22222") },
                { "Firebrick1", SKColor.Parse("ff3030") },
                { "Firebrick2", SKColor.Parse("ee2c2c") },
                { "Firebrick3", SKColor.Parse("cd2626") },
                { "Firebrick4", SKColor.Parse("8b1a1a") },
                { "FloralWhite", SKColor.Parse("fffaf0") },
                { "ForestGreen", SKColor.Parse("228b22") },
                { "Gainsboro", SKColor.Parse("dcdcdc") },
                { "GhostWhite", SKColor.Parse("f8f8ff") },
                { "Gold1", SKColor.Parse("ffd700") },
                { "Gold2", SKColor.Parse("eec900") },
                { "Gold3", SKColor.Parse("cdad00") },
                { "Gold4", SKColor.Parse("8b7500") },
                { "Goldenrod", SKColor.Parse("daa520") },
                { "Goldenrod1", SKColor.Parse("ffc125") },
                { "Goldenrod2", SKColor.Parse("eeb422") },
                { "Goldenrod3", SKColor.Parse("cd9b1d") },
                { "Goldenrod4", SKColor.Parse("8b6914") },
                { "Gray", SKColor.Parse("bebebe") },
                { "Gray1", SKColor.Parse("030303") },
                { "Gray2", SKColor.Parse("050505") },
                { "Gray3", SKColor.Parse("080808") },
                { "Gray4", SKColor.Parse("0a0a0a") },
                { "Gray5", SKColor.Parse("0d0d0d") },
                { "Gray6", SKColor.Parse("0f0f0f") },
                { "Gray7", SKColor.Parse("121212") },
                { "Gray8", SKColor.Parse("141414") },
                { "Gray9", SKColor.Parse("171717") },
                { "Gray10", SKColor.Parse("1a1a1a") },
                { "Gray11", SKColor.Parse("1c1c1c") },
                { "Gray12", SKColor.Parse("1f1f1f") },
                { "Gray13", SKColor.Parse("212121") },
                { "Gray14", SKColor.Parse("242424") },
                { "Gray15", SKColor.Parse("262626") },
                { "Gray16", SKColor.Parse("292929") },
                { "Gray17", SKColor.Parse("2b2b2b") },
                { "Gray18", SKColor.Parse("2e2e2e") },
                { "Gray19", SKColor.Parse("303030") },
                { "Gray20", SKColor.Parse("333333") },
                { "Gray21", SKColor.Parse("363636") },
                { "Gray22", SKColor.Parse("383838") },
                { "Gray23", SKColor.Parse("3b3b3b") },
                { "Gray24", SKColor.Parse("3d3d3d") },
                { "Gray25", SKColor.Parse("404040") },
                { "Gray26", SKColor.Parse("424242") },
                { "Gray27", SKColor.Parse("454545") },
                { "Gray28", SKColor.Parse("474747") },
                { "Gray29", SKColor.Parse("4a4a4a") },
                { "Gray30", SKColor.Parse("4d4d4d") },
                { "Gray31", SKColor.Parse("4f4f4f") },
                { "Gray32", SKColor.Parse("525252") },
                { "Gray33", SKColor.Parse("545454") },
                { "Gray34", SKColor.Parse("575757") },
                { "Gray35", SKColor.Parse("595959") },
                { "Gray36", SKColor.Parse("5c5c5c") },
                { "Gray37", SKColor.Parse("5e5e5e") },
                { "Gray38", SKColor.Parse("616161") },
                { "Gray39", SKColor.Parse("636363") },
                { "Gray40", SKColor.Parse("666666") },
                { "Gray41", SKColor.Parse("696969") },
                { "Gray42", SKColor.Parse("6b6b6b") },
                { "Gray43", SKColor.Parse("6e6e6e") },
                { "Gray44", SKColor.Parse("707070") },
                { "Gray45", SKColor.Parse("737373") },
                { "Gray46", SKColor.Parse("757575") },
                { "Gray47", SKColor.Parse("787878") },
                { "Gray48", SKColor.Parse("7a7a7a") },
                { "Gray49", SKColor.Parse("7d7d7d") },
                { "Gray50", SKColor.Parse("7f7f7f") },
                { "Gray51", SKColor.Parse("828282") },
                { "Gray52", SKColor.Parse("858585") },
                { "Gray53", SKColor.Parse("878787") },
                { "Gray54", SKColor.Parse("8a8a8a") },
                { "Gray55", SKColor.Parse("8c8c8c") },
                { "Gray56", SKColor.Parse("8f8f8f") },
                { "Gray57", SKColor.Parse("919191") },
                { "Gray58", SKColor.Parse("949494") },
                { "Gray59", SKColor.Parse("969696") },
                { "Gray60", SKColor.Parse("999999") },
                { "Gray61", SKColor.Parse("9c9c9c") },
                { "Gray62", SKColor.Parse("9e9e9e") },
                { "Gray63", SKColor.Parse("a1a1a1") },
                { "Gray64", SKColor.Parse("a3a3a3") },
                { "Gray65", SKColor.Parse("a6a6a6") },
                { "Gray66", SKColor.Parse("a8a8a8") },
                { "Gray67", SKColor.Parse("ababab") },
                { "Gray68", SKColor.Parse("adadad") },
                { "Gray69", SKColor.Parse("b0b0b0") },
                { "Gray70", SKColor.Parse("b3b3b3") },
                { "Gray71", SKColor.Parse("b5b5b5") },
                { "Gray72", SKColor.Parse("b8b8b8") },
                { "Gray73", SKColor.Parse("bababa") },
                { "Gray74", SKColor.Parse("bdbdbd") },
                { "Gray75", SKColor.Parse("bfbfbf") },
                { "Gray76", SKColor.Parse("c2c2c2") },
                { "Gray77", SKColor.Parse("c4c4c4") },
                { "Gray78", SKColor.Parse("c7c7c7") },
                { "Gray79", SKColor.Parse("c9c9c9") },
                { "Gray80", SKColor.Parse("cccccc") },
                { "Gray81", SKColor.Parse("cfcfcf") },
                { "Gray82", SKColor.Parse("d1d1d1") },
                { "Gray83", SKColor.Parse("d4d4d4") },
                { "Gray84", SKColor.Parse("d6d6d6") },
                { "Gray85", SKColor.Parse("d9d9d9") },
                { "Gray86", SKColor.Parse("dbdbdb") },
                { "Gray87", SKColor.Parse("dedede") },
                { "Gray88", SKColor.Parse("e0e0e0") },
                { "Gray89", SKColor.Parse("e3e3e3") },
                { "Gray90", SKColor.Parse("e5e5e5") },
                { "Gray91", SKColor.Parse("e8e8e8") },
                { "Gray92", SKColor.Parse("ebebeb") },
                { "Gray93", SKColor.Parse("ededed") },
                { "Gray94", SKColor.Parse("f0f0f0") },
                { "Gray95", SKColor.Parse("f2f2f2") },
                { "Gray97", SKColor.Parse("f7f7f7") },
                { "Gray98", SKColor.Parse("fafafa") },
                { "Gray99", SKColor.Parse("fcfcfc") },
                { "Green1", SKColor.Parse("00ff00") },
                { "Green2", SKColor.Parse("00ee00") },
                { "Green3", SKColor.Parse("00cd00") },
                { "Green4", SKColor.Parse("008b00") },
                { "GreenYellow", SKColor.Parse("adff2f") },
                { "Honeydew1", SKColor.Parse("f0fff0") },
                { "Honeydew2", SKColor.Parse("e0eee0") },
                { "Honeydew3", SKColor.Parse("c1cdc1") },
                { "Honeydew4", SKColor.Parse("838b83") },
                { "HotPink", SKColor.Parse("ff69b4") },
                { "HotPink1", SKColor.Parse("ff6eb4") },
                { "HotPink2", SKColor.Parse("ee6aa7") },
                { "HotPink3", SKColor.Parse("cd6090") },
                { "HotPink4", SKColor.Parse("8b3a62") },
                { "IndianRed", SKColor.Parse("cd5c5c") },
                { "IndianRed1", SKColor.Parse("ff6a6a") },
                { "IndianRed2", SKColor.Parse("ee6363") },
                { "IndianRed3", SKColor.Parse("cd5555") },
                { "IndianRed4", SKColor.Parse("8b3a3a") },
                { "Ivory1", SKColor.Parse("fffff0") },
                { "Ivory2", SKColor.Parse("eeeee0") },
                { "Ivory3", SKColor.Parse("cdcdc1") },
                { "Ivory4", SKColor.Parse("8b8b83") },
                { "Khaki", SKColor.Parse("f0e68c") },
                { "Khaki1", SKColor.Parse("fff68f") },
                { "Khaki2", SKColor.Parse("eee685") },
                { "Khaki3", SKColor.Parse("cdc673") },
                { "Khaki4", SKColor.Parse("8b864e") },
                { "Lavender", SKColor.Parse("e6e6fa") },
                { "LavenderBlush1", SKColor.Parse("fff0f5") },
                { "LavenderBlush2", SKColor.Parse("eee0e5") },
                { "LavenderBlush3", SKColor.Parse("cdc1c5") },
                { "LavenderBlush4", SKColor.Parse("8b8386") },
                { "Lawngreen", SKColor.Parse("7cfc00") },
                { "LemonChiffon1", SKColor.Parse("fffacd") },
                { "LemonChiffon2", SKColor.Parse("eee9bf") },
                { "LemonChiffon3", SKColor.Parse("cdc9a5") },
                { "LemonChiffon4", SKColor.Parse("8b8970") },
                { "Light", SKColor.Parse("eedd82") },
                { "LightBlue", SKColor.Parse("add8e6") },
                { "LightBlue1", SKColor.Parse("bfefff") },
                { "LightBlue2", SKColor.Parse("b2dfee") },
                { "LightBlue3", SKColor.Parse("9ac0cd") },
                { "LightBlue4", SKColor.Parse("68838b") },
                { "LightCoral", SKColor.Parse("f08080") },
                { "LightCyan1", SKColor.Parse("e0ffff") },
                { "LightCyan2", SKColor.Parse("d1eeee") },
                { "LightCyan3", SKColor.Parse("b4cdcd") },
                { "LightCyan4", SKColor.Parse("7a8b8b") },
                { "LightGoldenrod1", SKColor.Parse("ffec8b") },
                { "LightGoldenrod2", SKColor.Parse("eedc82") },
                { "LightGoldenrod3", SKColor.Parse("cdbe70") },
                { "LightGoldenrod4", SKColor.Parse("8b814c") },
                { "LightGoldenrodYellow", SKColor.Parse("fafad2") },
                { "LightGray", SKColor.Parse("d3d3d3") },
                { "LightPink", SKColor.Parse("ffb6c1") },
                { "LightPink1", SKColor.Parse("ffaeb9") },
                { "LightPink2", SKColor.Parse("eea2ad") },
                { "LightPink3", SKColor.Parse("cd8c95") },
                { "LightPink4", SKColor.Parse("8b5f65") },
                { "LightSalmon1", SKColor.Parse("ffa07a") },
                { "LightSalmon2", SKColor.Parse("ee9572") },
                { "LightSalmon3", SKColor.Parse("cd8162") },
                { "LightSalmon4", SKColor.Parse("8b5742") },
                { "LightSeaGreen", SKColor.Parse("20b2aa") },
                { "LightSkyBlue", SKColor.Parse("87cefa") },
                { "LightSkyBlue1", SKColor.Parse("b0e2ff") },
                { "LightSkyBlue2", SKColor.Parse("a4d3ee") },
                { "LightSkyBlue3", SKColor.Parse("8db6cd") },
                { "LightSkyBlue4", SKColor.Parse("607b8b") },
                { "LightSlateBlue", SKColor.Parse("8470ff") },
                { "LightSlateGray", SKColor.Parse("778899") },
                { "LightSteelBlue", SKColor.Parse("b0c4de") },
                { "LightSteelBlue1", SKColor.Parse("cae1ff") },
                { "LightSteelBlue2", SKColor.Parse("bcd2ee") },
                { "LightSteelBlue3", SKColor.Parse("a2b5cd") },
                { "LightSteelBlue4", SKColor.Parse("6e7b8b") },
                { "LightYellow1", SKColor.Parse("ffffe0") },
                { "LightYellow2", SKColor.Parse("eeeed1") },
                { "LightYellow3", SKColor.Parse("cdcdb4") },
                { "LightYellow4", SKColor.Parse("8b8b7a") },
                { "LimeGreen", SKColor.Parse("32cd32") },
                { "Linen", SKColor.Parse("faf0e6") },
                { "Magenta", SKColor.Parse("ff00ff") },
                { "Magenta2", SKColor.Parse("ee00ee") },
                { "Magenta3", SKColor.Parse("cd00cd") },
                { "Magenta4", SKColor.Parse("8b008b") },
                { "Maroon", SKColor.Parse("b03060") },
                { "Maroon1", SKColor.Parse("ff34b3") },
                { "Maroon2", SKColor.Parse("ee30a7") },
                { "Maroon3", SKColor.Parse("cd2990") },
                { "Maroon4", SKColor.Parse("8b1c62") },
                { "Medium", SKColor.Parse("66cdaa") },
                { "MediumAquamarine", SKColor.Parse("66cdaa") },
                { "MediumBlue", SKColor.Parse("0000cd") },
                { "MediumOrchid", SKColor.Parse("ba55d3") },
                { "MediumOrchid1", SKColor.Parse("e066ff") },
                { "MediumOrchid2", SKColor.Parse("d15fee") },
                { "MediumOrchid3", SKColor.Parse("b452cd") },
                { "MediumOrchid4", SKColor.Parse("7a378b") },
                { "MediumPurple", SKColor.Parse("9370db") },
                { "MediumPurple1", SKColor.Parse("ab82ff") },
                { "MediumPurple2", SKColor.Parse("9f79ee") },
                { "MediumPurple3", SKColor.Parse("8968cd") },
                { "MediumPurple4", SKColor.Parse("5d478b") },
                { "MediumSeaGreen", SKColor.Parse("3cb371") },
                { "MediumSlateBlue", SKColor.Parse("7b68ee") },
                { "MediumSpringGreen", SKColor.Parse("00fa9a") },
                { "MediumTurquoise", SKColor.Parse("48d1cc") },
                { "MediumVioletRed", SKColor.Parse("c71585") },
                { "MidnightBlue", SKColor.Parse("191970") },
                { "MintCream", SKColor.Parse("f5fffa") },
                { "MistyRose1", SKColor.Parse("ffe4e1") },
                { "MistyRose2", SKColor.Parse("eed5d2") },
                { "MistyRose3", SKColor.Parse("cdb7b5") },
                { "MistyRose4", SKColor.Parse("8b7d7b") },
                { "Moccasin", SKColor.Parse("ffe4b5") },
                { "NavyBlue", SKColor.Parse("000080") },
                { "OldLace", SKColor.Parse("fdf5e6") },
                { "OliveDrab", SKColor.Parse("6b8e23") },
                { "OliveDrab1", SKColor.Parse("c0ff3e") },
                { "OliveDrab2", SKColor.Parse("b3ee3a") },
                { "OliveDrab4", SKColor.Parse("698b22") },
                { "Orange1", SKColor.Parse("ffa500") },
                { "Orange2", SKColor.Parse("ee9a00") },
                { "Orange3", SKColor.Parse("cd8500") },
                { "Orange4", SKColor.Parse("8b5a00") },
                { "OrangeRed1", SKColor.Parse("ff4500") },
                { "OrangeRed2", SKColor.Parse("ee4000") },
                { "OrangeRed3", SKColor.Parse("cd3700") },
                { "OrangeRed4", SKColor.Parse("8b2500") },
                { "OrangeWhite1", SKColor.Parse("ffdead") },
                { "OrangeWhite2", SKColor.Parse("eecfa1") },
                { "OrangeWhite3", SKColor.Parse("cdb38b") },
                { "OrangeWhite4", SKColor.Parse("8b795e") },
                { "Orchid", SKColor.Parse("da70d6") },
                { "Orchid1", SKColor.Parse("ff83fa") },
                { "Orchid2", SKColor.Parse("ee7ae9") },
                { "Orchid3", SKColor.Parse("cd69c9") },
                { "Orchid4", SKColor.Parse("8b4789") },
                { "Pale", SKColor.Parse("db7093") },
                { "PaleGoldenrod", SKColor.Parse("eee8aa") },
                { "PaleGreen", SKColor.Parse("98fb98") },
                { "PaleGreen1", SKColor.Parse("9aff9a") },
                { "PaleGreen2", SKColor.Parse("90ee90") },
                { "PaleGreen3", SKColor.Parse("7ccd7c") },
                { "PaleGreen4", SKColor.Parse("548b54") },
                { "PaleTurquoise", SKColor.Parse("afeeee") },
                { "PaleTurquoise1", SKColor.Parse("bbffff") },
                { "PaleTurquoise2", SKColor.Parse("aeeeee") },
                { "PaleTurquoise3", SKColor.Parse("96cdcd") },
                { "PaleTurquoise4", SKColor.Parse("668b8b") },
                { "PaleVioletRed", SKColor.Parse("db7093") },
                { "PaleVioletRed1", SKColor.Parse("ff82ab") },
                { "PaleVioletRed2", SKColor.Parse("ee799f") },
                { "PaleVioletRed3", SKColor.Parse("cd6889") },
                { "PaleVioletRed4", SKColor.Parse("8b475d") },
                { "PapayaWhip", SKColor.Parse("ffefd5") },
                { "PeachPuff1", SKColor.Parse("ffdab9") },
                { "PeachPuff2", SKColor.Parse("eecbad") },
                { "PeachPuff3", SKColor.Parse("cdaf95") },
                { "PeachPuff4", SKColor.Parse("8b7765") },
                { "Pink", SKColor.Parse("ffc0cb") },
                { "Pink1", SKColor.Parse("ffb5c5") },
                { "Pink2", SKColor.Parse("eea9b8") },
                { "Pink3", SKColor.Parse("cd919e") },
                { "Pink4", SKColor.Parse("8b636c") },
                { "Plum", SKColor.Parse("dda0dd") },
                { "Plum1", SKColor.Parse("ffbbff") },
                { "Plum2", SKColor.Parse("eeaeee") },
                { "Plum3", SKColor.Parse("cd96cd") },
                { "Plum4", SKColor.Parse("8b668b") },
                { "PowderBlue", SKColor.Parse("b0e0e6") },
                { "Purple", SKColor.Parse("a020f0") },
                { "Purple1", SKColor.Parse("9b30ff") },
                { "Purple2", SKColor.Parse("912cee") },
                { "Purple3", SKColor.Parse("7d26cd") },
                { "Purple4", SKColor.Parse("551a8b") },
                { "Red1", SKColor.Parse("ff0000") },
                { "Red2", SKColor.Parse("ee0000") },
                { "Red3", SKColor.Parse("cd0000") },
                { "Red4", SKColor.Parse("8b0000") },
                { "RosyBrown", SKColor.Parse("bc8f8f") },
                { "RosyBrown1", SKColor.Parse("ffc1c1") },
                { "RosyBrown2", SKColor.Parse("eeb4b4") },
                { "RosyBrown3", SKColor.Parse("cd9b9b") },
                { "RosyBrown4", SKColor.Parse("8b6969") },
                { "RoyalBlue", SKColor.Parse("4169e1") },
                { "RoyalBlue1", SKColor.Parse("4876ff") },
                { "RoyalBlue2", SKColor.Parse("436eee") },
                { "RoyalBlue3", SKColor.Parse("3a5fcd") },
                { "RoyalBlue4", SKColor.Parse("27408b") },
                { "SaddleBrown", SKColor.Parse("8b4513") },
                { "Salmon", SKColor.Parse("fa8072") },
                { "Salmon1", SKColor.Parse("ff8c69") },
                { "Salmon2", SKColor.Parse("ee8262") },
                { "Salmon3", SKColor.Parse("cd7054") },
                { "Salmon4", SKColor.Parse("8b4c39") },
                { "SandyBrown", SKColor.Parse("f4a460") },
                { "SeaGreen1", SKColor.Parse("54ff9f") },
                { "SeaGreen2", SKColor.Parse("4eee94") },
                { "SeaGreen3", SKColor.Parse("43cd80") },
                { "SeaGreen4", SKColor.Parse("2e8b57") },
                { "Seashell1", SKColor.Parse("fff5ee") },
                { "Seashell2", SKColor.Parse("eee5de") },
                { "Seashell3", SKColor.Parse("cdc5bf") },
                { "Seashell4", SKColor.Parse("8b8682") },
                { "Sienna", SKColor.Parse("a0522d") },
                { "Sienna1", SKColor.Parse("ff8247") },
                { "Sienna2", SKColor.Parse("ee7942") },
                { "Sienna3", SKColor.Parse("cd6839") },
                { "Sienna4", SKColor.Parse("8b4726") },
                { "SkyBlue", SKColor.Parse("87ceeb") },
                { "SkyBlue1", SKColor.Parse("87ceff") },
                { "SkyBlue2", SKColor.Parse("7ec0ee") },
                { "SkyBlue3", SKColor.Parse("6ca6cd") },
                { "SkyBlue4", SKColor.Parse("4a708b") },
                { "SlateBlue", SKColor.Parse("6a5acd") },
                { "SlateBlue1", SKColor.Parse("836fff") },
                { "SlateBlue2", SKColor.Parse("7a67ee") },
                { "SlateBlue3", SKColor.Parse("6959cd") },
                { "SlateBlue4", SKColor.Parse("473c8b") },
                { "SlateGray", SKColor.Parse("708090") },
                { "SlateGray1", SKColor.Parse("c6e2ff") },
                { "SlateGray2", SKColor.Parse("b9d3ee") },
                { "SlateGray3", SKColor.Parse("9fb6cd") },
                { "SlateGray4", SKColor.Parse("6c7b8b") },
                { "Snow1", SKColor.Parse("fffafa") },
                { "Snow2", SKColor.Parse("eee9e9") },
                { "Snow3", SKColor.Parse("cdc9c9") },
                { "Snow4", SKColor.Parse("8b8989") },
                { "SpringGreen1", SKColor.Parse("00ff7f") },
                { "SpringGreen2", SKColor.Parse("00ee76") },
                { "SpringGreen3", SKColor.Parse("00cd66") },
                { "SpringGreen4", SKColor.Parse("008b45") },
                { "SteelBlue", SKColor.Parse("4682b4") },
                { "SteelBlue1", SKColor.Parse("63b8ff") },
                { "SteelBlue2", SKColor.Parse("5cacee") },
                { "SteelBlue3", SKColor.Parse("4f94cd") },
                { "SteelBlue4", SKColor.Parse("36648b") },
                { "Tan", SKColor.Parse("d2b48c") },
                { "Tan1", SKColor.Parse("ffa54f") },
                { "Tan2", SKColor.Parse("ee9a49") },
                { "Tan3", SKColor.Parse("cd853f") },
                { "Tan4", SKColor.Parse("8b5a2b") },
                { "Thistle", SKColor.Parse("d8bfd8") },
                { "Thistle1", SKColor.Parse("ffe1ff") },
                { "Thistle2", SKColor.Parse("eed2ee") },
                { "Thistle3", SKColor.Parse("cdb5cd") },
                { "Thistle4", SKColor.Parse("8b7b8b") },
                { "Tomato1", SKColor.Parse("ff6347") },
                { "Tomato2", SKColor.Parse("ee5c42") },
                { "Tomato3", SKColor.Parse("cd4f39") },
                { "Tomato4", SKColor.Parse("8b3626") },
                { "Turquoise", SKColor.Parse("40e0d0") },
                { "Turquoise1", SKColor.Parse("00f5ff") },
                { "Turquoise2", SKColor.Parse("00e5ee") },
                { "Turquoise3", SKColor.Parse("00c5cd") },
                { "Turquoise4", SKColor.Parse("00868b") },
                { "Violet", SKColor.Parse("ee82ee") },
                { "VioletRed", SKColor.Parse("d02090") },
                { "VioletRed1", SKColor.Parse("ff3e96") },
                { "VioletRed2", SKColor.Parse("ee3a8c") },
                { "VioletRed3", SKColor.Parse("cd3278") },
                { "VioletRed4", SKColor.Parse("8b2252") },
                { "Wheat", SKColor.Parse("f5deb3") },
                { "Wheat1", SKColor.Parse("ffe7ba") },
                { "Wheat2", SKColor.Parse("eed8ae") },
                { "Wheat3", SKColor.Parse("cdba96") },
                { "Wheat4", SKColor.Parse("8b7e66") },
                { "White", SKColor.Parse("ffffff") },
                { "WhiteSmoke", SKColor.Parse("f5f5f5") },
                { "Yellow1", SKColor.Parse("ffff00") },
                { "Yellow2", SKColor.Parse("eeee00") },
                { "Yellow3", SKColor.Parse("cdcd00") },
                { "Yellow4", SKColor.Parse("8b8b00") },
                { "YellowGreen", SKColor.Parse("9acd32") }
            };

            foreach (FieldInfo field in typeof(SKColors).GetFields(BindingFlags.Static | BindingFlags.Public))
            {
                if (!_library.ContainsKey(field.Name))
                {
                    object? value = field.GetValue(null);
                    if (value is null)
                    {
                        continue;
                    }

                    _library[field.Name] = (SKColor)value;
                }
            }
        }
    }
}
