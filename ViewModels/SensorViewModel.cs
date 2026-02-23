using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using VUKVWeightApp.Domain;
using VUKVWeightApp.Services;
using VUKVWeightApp.Utils;

namespace VUKVWeightApp.ViewModels;

public sealed class SensorViewModel : ObservableObject
{
    private readonly SensorController _controller;
    private readonly string _name;
    private readonly ScaleId _scaleId;

    private string _ip;
    private bool _isConnected;
    private double _weightKg;
    private int _rawGross;
    private string _uvText = "";
    private DateTime _lastUpdate;

    public string Name => _name;

    public string Ip
    {
        get => _ip;
        set => Set(ref _ip, value);
    }

    public bool IsConnected
    {
        get => _isConnected;
        private set
        {
            if (Set(ref _isConnected, value))
            {
                ConnectCommand.RaiseCanExecuteChanged();
                DisconnectCommand.RaiseCanExecuteChanged();
                ZeroCommand.RaiseCanExecuteChanged();
            }
        }
    }

    public double WeightKg { get => _weightKg; private set => Set(ref _weightKg, value); }
    public int RawGross { get => _rawGross; private set => Set(ref _rawGross, value); }
    public string UvText { get => _uvText; private set => Set(ref _uvText, value); }

    public DateTime LastUpdate { get => _lastUpdate; private set => Set(ref _lastUpdate, value); }

    public AsyncRelayCommand ConnectCommand { get; }
    public RelayCommand DisconnectCommand { get; }
    public AsyncRelayCommand ZeroCommand { get; }

    public SensorViewModel(SensorController controller, DeviceEndpoint endpoint)
    {
        _controller = controller;
        _name = endpoint.Name;
        _scaleId = new ScaleId(endpoint.Name, 1);
        _ip = endpoint.Ip;

        ConnectCommand = new AsyncRelayCommand(ConnectAsync, () => !IsConnected);
        DisconnectCommand = new RelayCommand(Disconnect, () => IsConnected);
        ZeroCommand = new AsyncRelayCommand(ZeroAsync, () => IsConnected);

        _controller.Store.Updated += (_, __) => PullLatest();
    }

    private async Task ConnectAsync()
    {
        _controller.UpdateIp(_name, Ip);
        await _controller.ConnectAsync(_name, timeoutMs: 1500, ct: CancellationToken.None);
        IsConnected = true;
    }

    private void Disconnect()
    {
        _controller.Disconnect(_name);
        IsConnected = false;
    }

    private async Task ZeroAsync()
    {
        await _controller.ZeroAsync(_name, CancellationToken.None);
    }

    private void PullLatest()
    {
        var snap = _controller.Store.SnapshotScales();
        if (snap.TryGetValue(_scaleId, out var r))
        {
            WeightKg = r.WeightKg;
            RawGross = r.RawGross;
            UvText = r.UvByChannel is { Count: > 0 }
                ? string.Join("  ", r.UvByChannel.Select((v, i) => $"CH{i + 1}:{v}"))
                : "";
            LastUpdate = r.Timestamp;
        }
    }
}
