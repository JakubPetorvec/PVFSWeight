using System;
using System.Collections.Generic;

namespace VUKVWeightApp.Domain
{
    public enum Dgt4Mode
    {
        DepCh,
        Transm
    }

    public sealed record DeviceEndpoint(
        string Name,
        string Ip,
        int Port,
        byte UnitId,
        Dgt4Mode Mode,
        int PollIntervalMs
    );

    /// <summary>
    /// Device-level scale (gross) identifier. Each IP device is treated as one scale.
    /// </summary>
    public readonly record struct ScaleId(string DeviceName, int ScaleIndex);

    /// <summary>
    /// Individual load-cell / tensometer channel identifier inside a device.
    /// ChannelIndex is 1..4 for DEP.CH live page 3001.
    /// </summary>
    public readonly record struct TenzometerId(string DeviceName, int ChannelIndex);

    /// <summary>
    /// Gross reading of one device scale.
    /// </summary>
    public sealed record ScaleReading(
        ScaleId Id,
        DateTime Timestamp,
        double WeightKg,
        bool Stable,
        int RawGross,
        ushort InputStatus,
        IReadOnlyList<short> UvByChannel
    );

    /// <summary>
    /// Derived per-channel reading. WeightKg is computed from scale gross and µV ratio.
    /// </summary>
    public sealed record TenzometerReading(
        TenzometerId Id,
        DateTime Timestamp,
        short Uv,
        double WeightKg,
        double PercentOfDevice
    );

    public sealed record DeviceSnapshot(
        string DeviceName,
        DateTime Timestamp,
        IReadOnlyList<ScaleReading> Scales
    );

    public sealed record GroupDefinition(
        string Name,
        IReadOnlyList<TenzometerId> Members
    );

    public sealed record AggregatedSnapshot(
        DateTime Timestamp,
        IReadOnlyDictionary<ScaleId, ScaleReading> LatestScales,
        IReadOnlyDictionary<TenzometerId, TenzometerReading> LatestTenzometers,
        double TotalKg,
        IReadOnlyDictionary<string, double> GroupSumsKg
    );
}
