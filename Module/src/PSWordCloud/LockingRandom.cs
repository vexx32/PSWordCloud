using System;
using System.Collections.Generic;

namespace PSWordCloud
{
    class LockingRandom
    {
        private const int MinimumAngleCount = 4;
        private const int MaximumAngleCount = 14;

        private readonly Random _random;
        private readonly object _lock = new object();

        internal LockingRandom()
        {
            _random = new Random();
        }

        internal LockingRandom(int seed)
        {
            _random = new Random(seed);
        }

        internal float PickRandomQuadrant()
            => GetRandomInt(0, 4) * 90;

        internal float RandomFloat()
        {
            lock (_lock)
            {
                return (float)_random.NextDouble();
            }
        }

        internal float RandomFloat(float min, float max)
        {
            lock (_lock)
            {
                return _random.NextFloat(min, max);
            }
        }

        internal int GetRandomInt()
        {
            lock (_lock)
            {
                return _random.Next();
            }
        }

        internal int GetRandomInt(int min, int max)
        {
            lock (_lock)
            {
                return _random.Next(min, max);
            }
        }

        internal IList<T> Shuffle<T>(IList<T> items)
        {
            lock (_lock)
            {
                return _random.Shuffle(items);
            }
        }

        internal IList<float> GetRandomFloats(
            int min,
            int max,
            int minCount = MinimumAngleCount,
            int maxCount = MaximumAngleCount)
        {
            var angles = new float[GetRandomInt(minCount, maxCount)];
            for (var index = 0; index < angles.Length; index++)
            {
                angles[index] = RandomFloat(min, max);
            }

            return angles;
        }
    }
}
