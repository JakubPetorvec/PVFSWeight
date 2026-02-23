using System;
using System.Collections.Generic;
using System.Linq;

namespace VUKVWeightApp.Utils
{
    /// <summary>
    /// Jednoduchý klouzavý mediánový filtr.
    /// Používá pevnou velikost okna (typicky liché číslo: 5, 7, 9).
    /// </summary>
    public sealed class MedianFilter
    {
        private readonly int _size;
        private readonly Queue<double> _buffer;

        public MedianFilter(int size)
        {
            if (size <= 0)
                throw new ArgumentOutOfRangeException(nameof(size));

            _size = size;
            _buffer = new Queue<double>(size);
        }

        /// <summary>
        /// Přidá nový vzorek a vrátí medián aktuálního okna.
        /// </summary>
        public double Add(double value)
        {
            if (_buffer.Count == _size)
                _buffer.Dequeue();

            _buffer.Enqueue(value);

            // medián
            var ordered = _buffer.OrderBy(v => v).ToArray();
            int mid = ordered.Length / 2;

            if (ordered.Length % 2 == 1)
                return ordered[mid];

            return (ordered[mid - 1] + ordered[mid]) / 2.0;
        }

        /// <summary>
        /// Vyprázdní filtr.
        /// </summary>
        public void Clear()
        {
            _buffer.Clear();
        }
    }
}
