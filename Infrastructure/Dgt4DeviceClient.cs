using System;
using System.Threading;
using System.Threading.Tasks;
using VUKVWeightApp.Domain;

namespace VUKVWeightApp.Infrastructure;

/// <summary>
/// High-level client for a single DGT4 endpoint (one device / one IP).
/// Uses DEP.CH page 3001 and treats the device as a single scale (ScaleIndex=1).
/// </summary>
public sealed class Dgt4DeviceClient : IDisposable
{
    private readonly SemaphoreSlim _gate = new(1, 1);
    private readonly Dgt4ModbusClient _client = new();

    public DeviceEndpoint Endpoint { get; }
    public bool IsConnected { get; private set; }

    public Dgt4DeviceClient(DeviceEndpoint endpoint)
    {
        Endpoint = endpoint ?? throw new ArgumentNullException(nameof(endpoint));
        if (endpoint.Mode != Dgt4Mode.DepCh)
            throw new NotSupportedException("This refactor build currently supports only DEP.CH for per-IP devices.");
    }

    public async Task ConnectAsync(int timeoutMs, CancellationToken ct)
    {
        await _gate.WaitAsync(ct);
        try
        {
            if (IsConnected) return;
            await _client.ConnectAsync(Endpoint.Ip, Endpoint.Port, timeoutMs);
            // Make sure the device is on page 3001 for live readings
            _client.ChangePage(Endpoint.UnitId, 3001);
            IsConnected = true;
        }
        finally
        {
            _gate.Release();
        }
    }

    public void Disconnect()
    {
        _gate.Wait();
        try
        {
            _client.Dispose();
            IsConnected = false;
        }
        finally
        {
            _gate.Release();
        }
    }

    public async Task ZeroAsync(CancellationToken ct)
    {
        await _gate.WaitAsync(ct);
        try
        {
            if (!IsConnected) return;
            _client.Zero(Endpoint.UnitId);
        }
        finally
        {
            _gate.Release();
        }
    }

    public async Task<ScaleReading?> ReadScaleAsync(double kgPerDiv, CancellationToken ct)
    {
        await _gate.WaitAsync(ct);
        try
        {
            if (!IsConnected) return null;

            var live = _client.ReadLive3001_DepCh(Endpoint.UnitId);

            // NOTE: stable bit mapping differs by configuration.
            // Keep conservative default: stable=true (UI can still show value without flicker).
            bool stable = true;

            double kg = live.GrossRaw * kgPerDiv;

            var id = new ScaleId(Endpoint.Name, 1);
            var uvs = new short[] { live.Uv1, live.Uv2, live.Uv3, live.Uv4 };
            return new ScaleReading(
                Id: id,
                Timestamp: DateTime.Now,
                WeightKg: kg,
                Stable: stable,
                RawGross: live.GrossRaw,
                InputStatus: live.InputStatus,
                UvByChannel: uvs
            );
        }
        finally
        {
            _gate.Release();
        }
    }

    public void Dispose()
    {
        Disconnect();
        _gate.Dispose();
    }
}
