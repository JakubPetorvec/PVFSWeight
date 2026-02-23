using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using VUKVWeightApp.Domain;
using VUKVWeightApp.Infrastructure;

namespace VUKVWeightApp.Services;

/// <summary>
/// Composition root for two devices (two IPs) and polling.
/// </summary>
public sealed class SensorController : IDisposable
{
    public MeasurementStore Store { get; } = new();

    private readonly Dictionary<string, DeviceEndpoint> _endpoints = new();
    private readonly Dictionary<string, Dgt4DeviceClient> _devices = new();
    private readonly Dictionary<string, SensorPoller> _pollers = new();

    private readonly int _medianWindow;
    private readonly double _kgPerDiv;
    private readonly int _intervalMs;

    public IReadOnlyDictionary<string, Dgt4DeviceClient> Devices => _devices;

    public SensorController(IEnumerable<DeviceEndpoint> endpoints, int medianWindow, double kgPerDiv, int intervalMs)
    {
        _medianWindow = medianWindow;
        _kgPerDiv = kgPerDiv;
        _intervalMs = intervalMs;
        foreach (var ep in endpoints)
        {
            _endpoints[ep.Name] = ep;
            var d = new Dgt4DeviceClient(ep);
            _devices[ep.Name] = d;
            _pollers[ep.Name] = new SensorPoller(d, Store, medianWindow, kgPerDiv, intervalMs);
        }
    }

    public void UpdateIp(string name, string ip)
    {
        if (!_endpoints.TryGetValue(name, out var ep)) return;

        // Recreate client & poller so the new IP is used.
        Disconnect(name);

        if (_pollers.TryGetValue(name, out var oldP)) oldP.Dispose();
        if (_devices.TryGetValue(name, out var oldD)) oldD.Dispose();

        var newEp = ep with { Ip = ip };
        _endpoints[name] = newEp;

        var d = new Dgt4DeviceClient(newEp);
        _devices[name] = d;
        _pollers[name] = new SensorPoller(d, Store, _medianWindow, _kgPerDiv, _intervalMs);
    }

    public async Task ConnectAsync(string name, int timeoutMs, CancellationToken ct)
    {
        if (!_devices.TryGetValue(name, out var dev)) return;
        await dev.ConnectAsync(timeoutMs, ct);
        if (_pollers.TryGetValue(name, out var p)) p.Start();
    }

    public void Disconnect(string name)
    {
        if (_pollers.TryGetValue(name, out var p)) p.Stop();
        if (_devices.TryGetValue(name, out var dev)) dev.Disconnect();
    }

    public async Task ZeroAsync(string name, CancellationToken ct)
    {
        if (_devices.TryGetValue(name, out var dev)) await dev.ZeroAsync(ct);
    }

    public void Dispose()
    {
        foreach (var p in _pollers.Values) p.Dispose();
        foreach (var d in _devices.Values) d.Dispose();
    }
}
