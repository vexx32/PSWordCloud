using System.Collections.Generic;
using System.Globalization;

namespace PSWordCloud
{
    internal class Word
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
    }
}
