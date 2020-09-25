using System.Collections.Generic;

namespace PSWordCloud
{
    internal struct WordListStatistics
    {
        internal float AverageLength { get; }

        internal float AverageFrequency { get; }

        internal float TotalFrequency { get; }

        internal int TotalLength { get; }

        internal int Count { get; }

        internal WordListStatistics(IReadOnlyList<Word> list)
        {
            Count = list.Count;

            float totalFrequency = 0;
            int totalLength = 0;
            for (int i = 0; i < list.Count; i++)
            {
                totalFrequency += list[i].RelativeSize;
                totalLength += list[i].Text.Length;
            }

            AverageFrequency = totalFrequency / Count;
            AverageLength = totalLength / Count;
            TotalFrequency = totalFrequency;
            TotalLength = totalLength;
        }
    }
}
