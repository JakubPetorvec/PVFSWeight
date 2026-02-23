using System;
using System.Threading;
using System.Threading.Tasks;
using VUKVWeightApp.Domain;
using VUKVWeightApp.Infrastructure;

namespace VUKVWeightApp.Services;

public sealed class SensorPoller : IDisposable
{
    private readonly Dgt4DeviceClient _device;
    private readonly MeasurementStore _store;
    private readonly MedianFilter _medianScale;
    private readonly MedianFilter[] _medianChannels;

    public double KgPerDiv { get; set; }
    public int IntervalMs { get; set; }

    private CancellationTokenSource? _cts;
    private Task? _loop;

    public SensorPoller(Dgt4DeviceClient device, MeasurementStore store, int medianWindow, double kgPerDiv, int intervalMs)
    {
        _device = device ?? throw new ArgumentNullException(nameof(device));
        _store = store ?? throw new ArgumentNullException(nameof(store));
        _medianScale = new MedianFilter(medianWindow);
        _medianChannels = new[]
        {
            new MedianFilter(medianWindow),
            new MedianFilter(medianWindow),
            new MedianFilter(medianWindow),
            new MedianFilter(medianWindow)
        };
        KgPerDiv = kgPerDiv;
        IntervalMs = Math.Max(50, intervalMs);
    }

    public bool IsRunning => _loop != null && !_loop.IsCompleted;

    public void Start()
    {
        if (IsRunning) return;
        _cts = new CancellationTokenSource();
        _loop = Task.Run(() => LoopAsync(_cts.Token));
    }

    public void Stop()
    {
        try { _cts?.Cancel(); } catch { }
    }

    private async Task LoopAsync(CancellationToken ct)
    {
        while (!ct.IsCancellationRequested)
        {
            try
            {
                var r = await _device.ReadScaleAsync(KgPerDiv, ct);
                if (r != null)
                {
                    // Smooth the device gross weight
                    var smoothKg = _medianScale.Push(r.WeightKg);
                    var sm = r with { WeightKg = smoothKg };
                    _store.UpdateScale(sm);

                    // Derive per-channel weights from µV ratio.
                    // Heuristic: abs(uv) / sum(abs(uv)) * deviceGrossKg.
                    // If sum==0, percent=0 and weight=0.
                    var uvs = sm.UvByChannel;
                    double sumAbs = 0;
                    for (int i = 0; i < uvs.Count; i++) sumAbs += Math.Abs((double)uvs[i]);

                    for (int i = 0; i < uvs.Count && i < 4; i++)
                    {
                        double pct = (sumAbs <= 0.000001) ? 0.0 : Math.Abs((double)uvs[i]) / sumAbs;
                        double w = smoothKg * pct;
                        w = _medianChannels[i].Push(w);
                        var tid = new TenzometerId(sm.Id.DeviceName, i + 1);
                        _store.UpdateTenzometer(new TenzometerReading(
                            Id: tid,
                            Timestamp: sm.Timestamp,
                            Uv: uvs[i],
                            WeightKg: w,
                            PercentOfDevice: pct
                        ));
                    }
                }
            }
            catch
            {
                // Device read errors are tolerated; UI will show stale values.
            }

            try { await Task.Delay(IntervalMs, ct); } catch { }
        }
    }

    public void Dispose()
    {
        Stop();
        try { _loop?.Wait(250); } catch { }
        _cts?.Dispose();
    }
}
