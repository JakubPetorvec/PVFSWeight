using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using VUKVWeightApp.Domain;

namespace VUKVWeightApp.Services;

/// <summary>
/// Thread-safe storage of latest readings.
/// - Latest scale (device gross) per <see cref="ScaleId"/>
/// - Latest per-channel derived tensometer weight per <see cref="TenzometerId"/>
/// </summary>
public sealed class MeasurementStore
{
    private readonly ConcurrentDictionary<ScaleId, ScaleReading> _scales = new();
    private readonly ConcurrentDictionary<TenzometerId, TenzometerReading> _tenzos = new();

    public event EventHandler? Updated;

    public void UpdateScale(ScaleReading reading)
    {
        _scales[reading.Id] = reading;
        Updated?.Invoke(this, EventArgs.Empty);
    }

    public void UpdateTenzometer(TenzometerReading reading)
    {
        _tenzos[reading.Id] = reading;
        Updated?.Invoke(this, EventArgs.Empty);
    }

    public IReadOnlyDictionary<ScaleId, ScaleReading> SnapshotScales()
        => new Dictionary<ScaleId, ScaleReading>(_scales);

    public IReadOnlyDictionary<TenzometerId, TenzometerReading> SnapshotTenzometers()
        => new Dictionary<TenzometerId, TenzometerReading>(_tenzos);
}
