using VUKVWeightApp.Utils;

namespace VUKVWeightApp.ViewModels;

public sealed class GroupBarViewModel : ObservableObject
{
    private string _name = "";
    private double _weightKg;
    private double _percent;

    public string Name { get => _name; set => Set(ref _name, value); }

    public double WeightKg { get => _weightKg; set => Set(ref _weightKg, value); }

    /// <summary>0..100</summary>
    public double Percent { get => _percent; set => Set(ref _percent, value); }

    public string PercentText => $"{Percent:0.0}%";
    public string WeightText => $"{WeightKg:0.###} kg";
}
