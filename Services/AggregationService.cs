using System;
using System.Collections.Generic;
using System.Linq;
using VUKVWeightApp.Domain;

namespace VUKVWeightApp.Services;

public sealed class AggregationService
{
    private readonly MeasurementStore _store;
    private readonly IReadOnlyList<GroupDefinition> _groups;

    public event EventHandler<AggregatedSnapshot>? Aggregated;

    public AggregationService(MeasurementStore store, IEnumerable<GroupDefinition> groups)
    {
        _store = store ?? throw new ArgumentNullException(nameof(store));
        _groups = (groups ?? Array.Empty<GroupDefinition>()).ToList();

        _store.Updated += (_, __) => Recompute();
    }

    public void Recompute()
    {
        var latestScales = _store.SnapshotScales();
        if (latestScales.Count == 0) return;
        var latestTenzos = _store.SnapshotTenzometers();

        double total = latestScales.Values.Sum(v => v.WeightKg);

        var sums = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
        foreach (var g in _groups)
        {
            double s = 0;
            foreach (var id in g.Members)
                if (latestTenzos.TryGetValue(id, out var r)) s += r.WeightKg;
            sums[g.Name] = s;
        }

        var snap = new AggregatedSnapshot(
            Timestamp: DateTime.Now,
            LatestScales: latestScales,
            LatestTenzometers: latestTenzos,
            TotalKg: total,
            GroupSumsKg: sums
        );

        Aggregated?.Invoke(this, snap);
    }
}
