using System.Collections.Generic;
using System.Globalization;

namespace PSWordCloud
{
    internal class Word
    {
        internal Word(string text, float relativeSize)
        {
            Text = text;
            RelativeSize = relativeSize;
        }

        internal string Text { get; set; }

        internal float RelativeSize { get; set; }

        internal float ScaledSize { get; private set; }

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
