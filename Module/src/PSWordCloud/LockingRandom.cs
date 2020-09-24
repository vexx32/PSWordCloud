using System;
using System.Collections.Generic;

namespace PSWordCloud
{
    /// <summary>
    /// Defines the <see cref="LockingRandom"/> class, a threadsafe wrapper for <see cref="Random"/>.
    /// </summary>
    class LockingRandom
    {
        private const int MinimumAngleCount = 4;
        private const int MaximumAngleCount = 14;

        private readonly Random _random;
        private readonly object _lock = new object();

        /// <summary>
        /// Creates a new instance of the <see cref="LockingRandom"/> class.
        /// </summary>
        internal LockingRandom()
        {
            _random = new Random();
        }

        /// <summary>
        /// Creates a new instance of the <see cref="LockingRandom"/> class with a set <paramref name="seed"/> value.
        /// </summary>
        /// <param name="seed">The seed value to create the <see cref="Random"/> instance with.</param>
        internal LockingRandom(int seed)
        {
            _random = new Random(seed);
        }

        /// <summary>
        /// Returns a random starting angle in degrees from 0, 90, 180, or 270.
        /// </summary>
        internal float PickRandomQuadrant()
            => GetRandomInt(0, 4) * 90;

        /// <summary>
        /// Returns a random float value between 0 and 1.
        /// </summary>
        internal float RandomFloat()
        {
            lock (_lock)
            {
                return (float)_random.NextDouble();
            }
        }

        /// <summary>
        /// Returns a random <see cref="float"/> value <paramref name="min"/> and <paramref name="max"/>.
        /// </summary>
        /// <param name="min">The minimum permissible value.</param>
        /// <param name="max">The maximum permissible value.</param>
        internal float RandomFloat(float min, float max)
        {
            lock (_lock)
            {
                return _random.NextFloat(min, max);
            }
        }

        /// <summary>
        /// Returns a random <see cref="int"/> value.
        /// </summary>
        internal int GetRandomInt()
        {
            lock (_lock)
            {
                return _random.Next();
            }
        }

        /// <summary>
        /// Returns a random <see cref="int"/> value between <paramref name="min"/> and <paramref name="max"/>.
        /// </summary>
        /// <param name="min">The minimum permissible value.</param>
        /// <param name="max">The maximum permissible value.</param>
        internal int GetRandomInt(int min, int max)
        {
            lock (_lock)
            {
                return _random.Next(min, max);
            }
        }

        /// <summary>
        /// Shuffles the given list of <paramref name="items"/> in-place.
        /// </summary>
        /// <param name="items">The list of items to shuffle.</param>
        /// <typeparam name="T">The type of items in the list.</typeparam>
        /// <returns>A reference to the input list as <see cref="IList{T}"/></returns>
        internal IList<T> Shuffle<T>(IList<T> items)
        {
            lock (_lock)
            {
                return _random.Shuffle(items);
            }
        }

        /// <summary>
        /// Returns a random list of <see cref="float"/> values constrained by the given min and max values.
        /// </summary>
        /// <param name="min">The smallest acceptable value.</param>
        /// <param name="max">The largest acceptable value.</param>
        /// <param name="minCount">The smallest allowable number of values to return.</param>
        /// <param name="maxCount">The largest allowable number of values to return.</param>
        internal IList<float> GetRandomFloats(
            int min,
            int max,
            int minCount = MinimumAngleCount,
            int maxCount = MaximumAngleCount)
        {
            var values = new float[GetRandomInt(minCount, maxCount)];
            for (var index = 0; index < values.Length; index++)
            {
                values[index] = RandomFloat(min, max);
            }

            return values;
        }
    }
}
