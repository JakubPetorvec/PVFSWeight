using System;
using System.Collections.Generic;
using System.Linq;

namespace VUKVWeightApp.Services;

/// <summary>
/// Simple rolling median filter.
/// </summary>
public sealed class MedianFilter
{
    private readonly Queue<double> _buf;
    public int Window { get; }

    public MedianFilter(int window)
    {
        Window = Math.Max(1, window);
        _buf = new Queue<double>(Window);
    }

    public double Push(double value)
    {
        if (_buf.Count == Window) _buf.Dequeue();
        _buf.Enqueue(value);

        var arr = _buf.ToArray();
        Array.Sort(arr);
        int mid = arr.Length / 2;
        return arr.Length % 2 == 1
            ? arr[mid]
            : (arr[mid - 1] + arr[mid]) / 2.0;
    }

    public void Reset() => _buf.Clear();
}
