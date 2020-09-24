using System;
using SkiaSharp;

namespace PSWordCloud
{
    /// <summary>
    /// Defines the <see cref="Word"/> class, containing the actual string as well as size/scale information and the
    /// necessary paths required to draw the word in an image.
    /// </summary>
    internal class Word : IDisposable
    {
        /// <summary>
        /// Creates a new instance of the <see cref="Word"/> class.
        /// </summary>
        /// <param name="text">The string that comprises the word.</param>
        /// <param name="relativeSize">The size of this word relative to others in its collection.</param>
        internal Word(string text, float relativeSize)
            : this(text, relativeSize, isFocusWord: false)
        {
        }

        /// <summary>
        /// Creates a new instance of the <see cref="Word"/> class.
        /// </summary>
        /// <param name="text">The string that comprises the word.</param>
        /// <param name="relativeSize">The size of this word relative to others in its collection.</param>
        /// <param name="isFocusWord">If true, marks this word as the focus word for the collection.</param>
        internal Word(string text, float relativeSize, bool isFocusWord)
        {
            Text = text;
            RelativeSize = relativeSize;
            IsFocusWord = isFocusWord;
        }

        /// <summary>
        /// Gets the text of the word.
        /// </summary>
        internal string Text { get; }

        /// <summary>
        /// Gets the relative size of the word.
        /// </summary>
        internal float RelativeSize { get; }

        /// <summary>
        /// Gets the fully scaled size of the word.
        /// </summary>
        internal float ScaledSize { get; private set; }

        /// <summary>
        /// Gets a value which indicates whether the word should be prominently featured.
        /// </summary>
        internal bool IsFocusWord { get; } = false;

        /// <summary>
        /// Gets or sets the <see cref="SKPath"/> value which determines how it is drawn in the image.
        /// </summary>
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

        /// <summary>
        /// Gets or sets the <see cref="SKPath"/> that defines the bubble which will surround the word in the image.
        /// </summary>
        internal SKPath? Bubble { get; set; } = null;

        internal SKRect Bounds {
            get => _bounds ?? Bubble?.TightBounds ?? Path.TightBounds;
            set => _bounds = value;
        }

        /// <summary>
        /// Gets or sets the width of the padding space on each side of the word.
        /// </summary>
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

        /// <summary>
        /// Returns the <see cref="Text"/> of the word.
        /// </summary>
        public override string ToString() => Text;

    /// <summary>
    /// Disposes the <see cref="IDisposable"/> managed objects owned by this instance.
    /// </summary>
        public void Dispose()
        {
            _path?.Dispose();
            Bubble?.Dispose();
        }
    }
}
