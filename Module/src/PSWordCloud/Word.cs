using System;
using SkiaSharp;

namespace PSWordCloud
{
    internal class Word : IDisposable
    {
        internal Word(string text, float relativeSize)
            : this(text, relativeSize, isFocusWord: false)
        {
        }

        internal Word(string text, float relativeSize, bool isFocusWord)
        {
            Text = text;
            RelativeSize = relativeSize;
            IsFocusWord = isFocusWord;
        }

        internal string Text { get; set; }

        internal float RelativeSize { get; set; }

        internal float ScaledSize { get; private set; }

        internal bool IsFocusWord { get; private set; } = false;

        internal SKPath Path {
            get => _path ??= new SKPath();
            set
            {
                if (_path is not null)
                {
                    _path.Dispose();
                }

                _path = value;
            }
        }

        internal SKPath? Bubble { get; set; } = null;

        internal SKRect Bounds {
            get => _bounds ?? Bubble?.TightBounds ?? Path.TightBounds;
            set => _bounds = value;
        }

        internal float Padding { get; set; } = 0;

        private SKRect? _bounds = null;
        private SKPath? _path = null;

        /// <summary>
        /// Scale the <see cref="Word"/> by the global scale value and determine its final size.
        /// </summary>
        /// <param name="baseSize">The base size for the font.</param>
        /// <param name="globalScale">The global scaling factor.</param>
        /// <param name="allWords">The dictionary of word scales containing their base sizes.</param>
        /// <returns>The scaled word size.</returns>
        internal float Scale(float globalScale)
        {
            ScaledSize = RelativeSize * globalScale;
            return ScaledSize;
        }

        public override string ToString() => Text;

        public void Dispose()
        {
            _path?.Dispose();
            Bubble?.Dispose();
        }
    }
}
