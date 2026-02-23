using System.Globalization;
using VUKVWeightApp.Utils;

namespace VUKVWeightApp.ViewModels;

public sealed class ChannelRowViewModel : ObservableObject
{
    public ChannelRowViewModel(string channel, double mvPerV)
    {
        _channel = channel;
        _mvPerV = mvPerV;
    }

    public string Channel => _channel;
    private readonly string _channel;

    public short Uv
    {
        get => _uv;
        set => Set(ref _uv, value);
    }
    private short _uv;

    public double MvPerV
    {
        get => _mvPerV;
        set
        {
            if (Set(ref _mvPerV, value))
                Raise(nameof(MvPerV));
        }
    }
    private double _mvPerV;

    public double Kg
    {
        get => _kg;
        set
        {
            if (Set(ref _kg, value))
            {
                Raise(nameof(KgText));
            }
        }
    }
    private double _kg;

    public double Percent
    {
        get => _percent;
        set
        {
            if (Set(ref _percent, value))
            {
                Raise(nameof(PercentText));
            }
        }
    }
    private double _percent;

    public string KgText => Kg.ToString("0.00", CultureInfo.InvariantCulture);
    public string PercentText => Percent.ToString("0.0", CultureInfo.InvariantCulture);
}
