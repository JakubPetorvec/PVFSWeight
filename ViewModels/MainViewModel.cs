using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Win32;
using VUKVWeightApp.Domain;
using VUKVWeightApp.Infrastructure;
using VUKVWeightApp.Utils;

namespace VUKVWeightApp.ViewModels
{
    public sealed class MainViewModel : INotifyPropertyChanged
    {
        private const byte FixedUnitId = 1;

        private int _port = 502;
        private int _timeoutMs = 1500;
        private int _samplingMs = 200;
        private int _recordingMs = 250; // 4 Hz

        private double _uvDeadband = 0.0;
        private double _zeroTolKg = 0.0;

        // ✅ NE readonly – po Disconnect se klient vytvoří znovu
        private Dgt4ModbusClient _scale1 = new();
        private Dgt4ModbusClient _scale2 = new();

        private CancellationTokenSource? _cts;
        private Task? _loop1;
        private Task? _loop2;

        private double _scale1Kg;
        private double _scale2Kg;
        private DateTime? _scale1Rx;
        private DateTime? _scale2Rx;

        private readonly double[] _s1ChKg = new double[4];
        private readonly double[] _s2ChKg = new double[4];

        public ObservableCollection<ChannelRowViewModel> Scale1Channels { get; } = new();
        public ObservableCollection<ChannelRowViewModel> Scale2Channels { get; } = new();


        // ✅ NOVÉ: citlivost load cellu (mV/V) pro každý kanál
        // Používá se pro korekci rozdílných loadcellů při rozdělení grossKg do rohů:
        // load ~ |uV| / (mV/V)
        private readonly double[] _s1MvPerV = new double[4] { 2.0047, 2.0016, 2.0070, 2.0042 };
        private readonly double[] _s2MvPerV = new double[4] { 2.008, 2.0055, 2.0073, 2.0080 };

        private bool _isRecording;
        private CancellationTokenSource? _recCts;
        private Task? _recTask;

        private readonly (int scale, int ch)[] _g1 = new[] { (1, 1), (1, 2) };
        private readonly (int scale, int ch)[] _g2 = new[] { (1, 3), (1, 4) };
        private readonly (int scale, int ch)[] _g3 = new[] { (2, 1), (2, 2) };
        private readonly (int scale, int ch)[] _g4 = new[] { (2, 3), (2, 4) };

        public MainViewModel()
        {
            // ✅ Default IP – vždy se ukáže v textboxu
            _scale1Ip = "192.168.199.123";
            _scale2Ip = "192.168.199.124";

            ConnectScale1Command = new RelayCommand(() => _ = ConnectScale1(), () => CanConnectScale1);
            DisconnectScale1Command = new RelayCommand(DisconnectScale1, () => IsScale1Connected);

            ConnectScale2Command = new RelayCommand(() => _ = ConnectScale2(), () => CanConnectScale2);
            DisconnectScale2Command = new RelayCommand(DisconnectScale2, () => IsScale2Connected);

            ConnectBothCommand = new RelayCommand(() => _ = ConnectBoth(), () => CanConnectAny);
            DisconnectBothCommand = new RelayCommand(DisconnectBoth, () => CanDisconnectAny);

            ZeroScale1Command = new AsyncRelayCommand(ZeroScale1Async, () => IsScale1Connected);
            ZeroScale2Command = new AsyncRelayCommand(ZeroScale2Async, () => IsScale2Connected);
            ZeroBothCommand = new AsyncRelayCommand(ZeroBothAsync, () => CanZeroBoth);

            ToggleRecordingCommand = new RelayCommand(ToggleRecording, () => true); // ✅ vždy “klikatelné”
            ClearSamplesCommand = new RelayCommand(() =>
            {
                Samples.Clear();
                ExportSamplesCommand.RaiseCanExecuteChanged();
            }, () => true);

            ExportSamplesCommand = new RelayCommand(ExportSamples, () => Samples.Count > 0);

            ExitCommand = new RelayCommand(() => App.Current.Shutdown(), () => true);

            OnPropertyChanged(nameof(Scale1Ip));
            OnPropertyChanged(nameof(Scale2Ip));
            OnPropertyChanged(nameof(RecordingButtonText));

            UpdateCanConnect();

            Group1Label = "G1";
            Group2Label = "G2";
            Group3Label = "G3";
            Group4Label = "G4";

            // kanály (mV/V je editovatelné v UI)
            Scale1Channels.Add(new ChannelRowViewModel("CH1", _s1MvPerV[0]));
            Scale1Channels.Add(new ChannelRowViewModel("CH2", _s1MvPerV[1]));
            Scale1Channels.Add(new ChannelRowViewModel("CH3", _s1MvPerV[2]));
            Scale1Channels.Add(new ChannelRowViewModel("CH4", _s1MvPerV[3]));

            Scale2Channels.Add(new ChannelRowViewModel("CH1", _s2MvPerV[0]));
            Scale2Channels.Add(new ChannelRowViewModel("CH2", _s2MvPerV[1]));
            Scale2Channels.Add(new ChannelRowViewModel("CH3", _s2MvPerV[2]));
            Scale2Channels.Add(new ChannelRowViewModel("CH4", _s2MvPerV[3]));

            Samples.CollectionChanged += (_, __) =>
            {
                ExportSamplesCommand.RaiseCanExecuteChanged();
            };

            UpdateComputed();
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        public string Scale1Ip
        {
            get => _scale1Ip;
            set { _scale1Ip = value; OnPropertyChanged(); UpdateCanConnect(); }
        }
        private string _scale1Ip = "";

        public string Scale2Ip
        {
            get => _scale2Ip;
            set { _scale2Ip = value; OnPropertyChanged(); UpdateCanConnect(); }
        }
        private string _scale2Ip = "";

        // ✅ NOVÉ: nastavitelné mV/V (podle certifikátu každého load cellu)
        // Příklad: Output 2.0080 mV/V => nastav příslušný kanál na 2.0080
        public double Scale1Ch1MvPerV { get => _s1MvPerV[0]; set { _s1MvPerV[0] = value; OnPropertyChanged(); } }
        public double Scale1Ch2MvPerV { get => _s1MvPerV[1]; set { _s1MvPerV[1] = value; OnPropertyChanged(); } }
        public double Scale1Ch3MvPerV { get => _s1MvPerV[2]; set { _s1MvPerV[2] = value; OnPropertyChanged(); } }
        public double Scale1Ch4MvPerV { get => _s1MvPerV[3]; set { _s1MvPerV[3] = value; OnPropertyChanged(); } }

        public double Scale2Ch1MvPerV { get => _s2MvPerV[0]; set { _s2MvPerV[0] = value; OnPropertyChanged(); } }
        public double Scale2Ch2MvPerV { get => _s2MvPerV[1]; set { _s2MvPerV[1] = value; OnPropertyChanged(); } }
        public double Scale2Ch3MvPerV { get => _s2MvPerV[2]; set { _s2MvPerV[2] = value; OnPropertyChanged(); } }
        public double Scale2Ch4MvPerV { get => _s2MvPerV[3]; set { _s2MvPerV[3] = value; OnPropertyChanged(); } }

        public bool IsScale1Connected
        {
            get => _isScale1Connected;
            private set
            {
                _isScale1Connected = value;
                OnPropertyChanged();
                UpdateCanConnect();
            }
        }
        private bool _isScale1Connected;

        public bool IsScale2Connected
        {
            get => _isScale2Connected;
            private set
            {
                _isScale2Connected = value;
                OnPropertyChanged();
                UpdateCanConnect();
            }
        }
        private bool _isScale2Connected;

        public bool CanConnectScale1 => !IsScale1Connected && !string.IsNullOrWhiteSpace(Scale1Ip);
        public bool CanConnectScale2 => !IsScale2Connected && !string.IsNullOrWhiteSpace(Scale2Ip);

        public bool CanConnectAny => CanConnectScale1 || CanConnectScale2;
        public bool CanDisconnectAny => IsScale1Connected || IsScale2Connected;
        public bool CanZeroBoth => IsScale1Connected || IsScale2Connected;

        public string Scale1WeightText
        {
            get => _scale1WeightText;
            private set { _scale1WeightText = value; OnPropertyChanged(); }
        }
        private string _scale1WeightText = "--- kg";

        public string Scale2WeightText
        {
            get => _scale2WeightText;
            private set { _scale2WeightText = value; OnPropertyChanged(); }
        }
        private string _scale2WeightText = "--- kg";

        public string TotalWeightText
        {
            get => _totalWeightText;
            private set { _totalWeightText = value; OnPropertyChanged(); }
        }
        private string _totalWeightText = "Celkem: --- kg";

        public string Group1Label { get => _group1Label; private set { _group1Label = value; OnPropertyChanged(); } }
        private string _group1Label = "G1";
        public string Group2Label { get => _group2Label; private set { _group2Label = value; OnPropertyChanged(); } }
        private string _group2Label = "G2";
        public string Group3Label { get => _group3Label; private set { _group3Label = value; OnPropertyChanged(); } }
        private string _group3Label = "G3";
        public string Group4Label { get => _group4Label; private set { _group4Label = value; OnPropertyChanged(); } }
        private string _group4Label = "G4";

        public double Group1Percent { get => _g1p; private set { _g1p = value; OnPropertyChanged(); } }
        private double _g1p;
        public double Group2Percent { get => _g2p; private set { _g2p = value; OnPropertyChanged(); } }
        private double _g2p;
        public double Group3Percent { get => _g3p; private set { _g3p = value; OnPropertyChanged(); } }
        private double _g3p;
        public double Group4Percent { get => _g4p; private set { _g4p = value; OnPropertyChanged(); } }
        private double _g4p;

        public string Group1ValueText { get => _g1t; private set { _g1t = value; OnPropertyChanged(); } }
        private string _g1t = "--- kg";
        public string Group2ValueText { get => _g2t; private set { _g2t = value; OnPropertyChanged(); } }
        private string _g2t = "--- kg";
        public string Group3ValueText { get => _g3t; private set { _g3t = value; OnPropertyChanged(); } }
        private string _g3t = "--- kg";
        public string Group4ValueText { get => _g4t; private set { _g4t = value; OnPropertyChanged(); } }
        private string _g4t = "--- kg";

        public ObservableCollection<SampleRow> Samples { get; } = new();

        public string RecordingButtonText => _isRecording ? "Zastavit záznam" : "Spustit záznam";

        public RelayCommand ConnectScale1Command { get; }
        public RelayCommand DisconnectScale1Command { get; }
        public RelayCommand ConnectScale2Command { get; }
        public RelayCommand DisconnectScale2Command { get; }

        public RelayCommand ConnectBothCommand { get; }
        public RelayCommand DisconnectBothCommand { get; }

        public AsyncRelayCommand ZeroScale1Command { get; }
        public AsyncRelayCommand ZeroScale2Command { get; }
        public AsyncRelayCommand ZeroBothCommand { get; }

        public RelayCommand ToggleRecordingCommand { get; }
        public RelayCommand ClearSamplesCommand { get; }

        public RelayCommand ExitCommand { get; }
        public RelayCommand ExportSamplesCommand { get; }

        private async Task ConnectScale1()
        {
            try
            {
                await _scale1.ConnectAsync(Scale1Ip.Trim(), _port, _timeoutMs);
                IsScale1Connected = true;
                StartLoopsIfNeeded();
            }
            catch (Exception ex)
            {
                IsScale1Connected = false;
                MessageBox.Show("Nepodařilo se připojit k Váha 1:\n" + ex.Message, "Chyba připojení", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                UpdateCanConnect();
            }
        }

        private async Task ConnectScale2()
        {
            try
            {
                await _scale2.ConnectAsync(Scale2Ip.Trim(), _port, _timeoutMs);
                IsScale2Connected = true;
                StartLoopsIfNeeded();
            }
            catch (Exception ex)
            {
                IsScale2Connected = false;
                MessageBox.Show("Nepodařilo se připojit k Váha 2:\n" + ex.Message, "Chyba připojení", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                UpdateCanConnect();
            }
        }

        private void DisconnectScale1()
        {
            try { _scale1.Dispose(); } catch { }
            _scale1 = new Dgt4ModbusClient(); // ✅ důležité

            IsScale1Connected = false;

            _scale1Kg = 0;
            Array.Clear(_s1ChKg, 0, 4);
            _scale1Rx = null;

            Scale1WeightText = "--- kg";

            UpdateComputed();
            StopLoopsIfNotNeeded();
        }

        private void DisconnectScale2()
        {
            try { _scale2.Dispose(); } catch { }
            _scale2 = new Dgt4ModbusClient(); // ✅ důležité

            IsScale2Connected = false;

            _scale2Kg = 0;
            Array.Clear(_s2ChKg, 0, 4);
            _scale2Rx = null;

            Scale2WeightText = "--- kg";

            UpdateComputed();
            StopLoopsIfNotNeeded();
        }

        private async Task ConnectBoth()
        {
            if (CanConnectScale1) await ConnectScale1();
            if (CanConnectScale2) await ConnectScale2();
        }

        private void DisconnectBoth()
        {
            if (IsScale1Connected) DisconnectScale1();
            if (IsScale2Connected) DisconnectScale2();
        }

        // ✅ U tebe je metoda Zero(...)
        private Task ZeroScale1Async() => Task.Run(() => { try { _scale1.Zero(FixedUnitId); } catch { } });
        private Task ZeroScale2Async() => Task.Run(() => { try { _scale2.Zero(FixedUnitId); } catch { } });

        private Task ZeroBothAsync()
        {
            return Task.Run(() =>
            {
                try { if (IsScale1Connected) _scale1.Zero(FixedUnitId); } catch { }
                try { if (IsScale2Connected) _scale2.Zero(FixedUnitId); } catch { }
            });
        }

        private void ToggleRecording()
        {
            // ✅ když není nic připojené, jen nic nedělej (ale tlačítko má text!)
            if (!(IsScale1Connected || IsScale2Connected))
                return;

            _isRecording = !_isRecording;
            OnPropertyChanged(nameof(RecordingButtonText));

            if (_isRecording) StartRecordingLoop();
            else StopRecordingLoop();
        }

        private void StartRecordingLoop()
        {
            if (_recCts != null)
                return;

            _recCts = new CancellationTokenSource();
            _recTask = Task.Run(() => RecordingLoop(_recCts.Token));
        }

        private void StopRecordingLoop()
        {
            if (_recCts == null)
                return;

            try { _recCts.Cancel(); } catch { }
            _recCts = null;
            _recTask = null;
        }

        private async Task RecordingLoop(CancellationToken ct)
        {
            while (!ct.IsCancellationRequested)
            {
                if (_isRecording && (IsScale1Connected || IsScale2Connected))
                {
                    try { AddSampleRow(); } catch { }
                }

                try { await Task.Delay(_recordingMs, ct); }
                catch { }
            }
        }

        private void StartLoopsIfNeeded()
        {
            if (_cts != null)
                return;

            _cts = new CancellationTokenSource();
            _loop1 = Task.Run(() => PollLoop(1, _cts.Token));
            _loop2 = Task.Run(() => PollLoop(2, _cts.Token));
        }

        private void StopLoopsIfNotNeeded()
        {
            if (_cts == null)
                return;

            if (IsScale1Connected || IsScale2Connected)
                return;

            try { _cts.Cancel(); } catch { }
            _cts = null;
            _loop1 = null;
            _loop2 = null;

            StopRecordingLoop();
        }

        private async Task PollLoop(int which, CancellationToken ct)
        {
            while (!ct.IsCancellationRequested)
            {
                try
                {
                    if (which == 1)
                    {
                        if (IsScale1Connected) ReadOne(1, _scale1);
                    }
                    else
                    {
                        if (IsScale2Connected) ReadOne(2, _scale2);
                    }
                }
                catch { }

                try { await Task.Delay(_samplingMs, ct); }
                catch { }
            }
        }

        private void ReadOne(int which, Dgt4ModbusClient client)
        {
            var live = client.ReadLive3001_DepCh(FixedUnitId);

            double grossKg = live.GrossRaw; // bez dalšího násobení (váha už posílá správně)
            if (_zeroTolKg > 0 && Math.Abs(grossKg) < _zeroTolKg) grossKg = 0;

            double[] chKg = which == 1 ? _s1ChKg : _s2ChKg;

            double u1 = ApplyUvDeadband(live.Uv1);
            double u2 = ApplyUvDeadband(live.Uv2);
            double u3 = ApplyUvDeadband(live.Uv3);
            double u4 = ApplyUvDeadband(live.Uv4);

            // ✅ GAME CHANGER: rozdělení podle |uV| korigované citlivostí mV/V:
            // load ~ |uV| / (mV/V)
            double n1 = Math.Abs(u1) / GetChannelMvPerV(which, 1);
            double n2 = Math.Abs(u2) / GetChannelMvPerV(which, 2);
            double n3 = Math.Abs(u3) / GetChannelMvPerV(which, 3);
            double n4 = Math.Abs(u4) / GetChannelMvPerV(which, 4);

            double sum = n1 + n2 + n3 + n4;

            if (sum <= 0.0001 || grossKg == 0)
            {
                chKg[0] = chKg[1] = chKg[2] = chKg[3] = 0;
            }
            else
            {
                chKg[0] = grossKg * (n1 / sum);
                chKg[1] = grossKg * (n2 / sum);
                chKg[2] = grossKg * (n3 / sum);
                chKg[3] = grossKg * (n4 / sum);
            }

            // props pro UI tenzometrů
            var chVm = which == 1 ? Scale1Channels : Scale2Channels;
            chVm[0].Uv = (short)u1; chVm[1].Uv = (short)u2; chVm[2].Uv = (short)u3; chVm[3].Uv = (short)u4;
            chVm[0].Kg = chKg[0]; chVm[1].Kg = chKg[1]; chVm[2].Kg = chKg[2]; chVm[3].Kg = chKg[3];

            double denom = Math.Abs(grossKg) < 0.0001 ? 0 : Math.Abs(grossKg);
            chVm[0].Percent = denom <= 0 ? 0 : (Math.Abs(chKg[0]) / denom) * 100.0;
            chVm[1].Percent = denom <= 0 ? 0 : (Math.Abs(chKg[1]) / denom) * 100.0;
            chVm[2].Percent = denom <= 0 ? 0 : (Math.Abs(chKg[2]) / denom) * 100.0;
            chVm[3].Percent = denom <= 0 ? 0 : (Math.Abs(chKg[3]) / denom) * 100.0;

            var now = DateTime.Now;

            if (which == 1)
            {
                _scale1Kg = grossKg;
                _scale1Rx = now;
                Scale1WeightText = FormatKg(grossKg);
            }
            else
            {
                _scale2Kg = grossKg;
                _scale2Rx = now;
                Scale2WeightText = FormatKg(grossKg);
            }

            UpdateComputed();
        }

        private double ApplyUvDeadband(double uv)
        {
            if (_uvDeadband <= 0) return uv;
            return Math.Abs(uv) < _uvDeadband ? 0 : uv;
        }

        private static double SafeSensitivity(double mvPerV)
        {
            // když si omylem nastavíš 0, nechceme dělení nulou
            return mvPerV > 0.000001 ? mvPerV : 1.0;
        }

        private double GetChannelMvPerV(int whichScale, int channelIndex1Based)
        {
            int i = channelIndex1Based - 1;
            if (i < 0 || i > 3) return 1.0;

            return whichScale == 1
                ? SafeSensitivity(Scale1Channels[i].MvPerV)
                : SafeSensitivity(Scale2Channels[i].MvPerV);
        }

        
        private void ExportSamples()
        {
            try
            {
                var dlg = new SaveFileDialog
                {
                    Title = "Export záznamů do Excelu",
                    Filter = "Excel (*.xlsx)|*.xlsx",
                    FileName = $"pvfs_zaznam_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
                };

                if (dlg.ShowDialog() != true)
                    return;

                ExcelExporter.ExportSamples(dlg.FileName, Samples);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Export do Excelu se nezdařil:\n" + ex.Message,
                    "Export",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

private void AddSampleRow()
        {
            string t1 = _scale1Rx.HasValue ? _scale1Rx.Value.ToString("HH:mm:ss.fff") : "";
            string t2 = _scale2Rx.HasValue ? _scale2Rx.Value.ToString("HH:mm:ss.fff") : "";

            double total = _scale1Kg + _scale2Kg;

            double g1 = GroupSum(_g1);
            double g2 = GroupSum(_g2);
            double g3 = GroupSum(_g3);
            double g4 = GroupSum(_g4);

            if (Samples.Count > 5000)
                Samples.RemoveAt(0);

            App.Current.Dispatcher.Invoke(() =>
            {
                Samples.Add(new SampleRow
                {
                    Scale1RxTime = t1,
                    Scale2RxTime = t2,
                    TotalKg = FormatKg(total),
                    Scale1Kg = FormatKg(_scale1Kg),
                    Scale2Kg = FormatKg(_scale2Kg),
                    G1Kg = FormatKg(g1),
                    G2Kg = FormatKg(g2),
                    G3Kg = FormatKg(g3),
                    G4Kg = FormatKg(g4),
                });
            });
        }

        private void UpdateComputed()
        {
            double total = _scale1Kg + _scale2Kg;
            TotalWeightText = "Celkem: " + FormatKg(total);

            double g1 = GroupSum(_g1);
            double g2 = GroupSum(_g2);
            double g3 = GroupSum(_g3);
            double g4 = GroupSum(_g4);

            Group1ValueText = FormatKg(g1);
            Group2ValueText = FormatKg(g2);
            Group3ValueText = FormatKg(g3);
            Group4ValueText = FormatKg(g4);

            Group1Percent = total <= 0.0001 ? 0 : (g1 / total) * 100.0;
            Group2Percent = total <= 0.0001 ? 0 : (g2 / total) * 100.0;
            Group3Percent = total <= 0.0001 ? 0 : (g3 / total) * 100.0;
            Group4Percent = total <= 0.0001 ? 0 : (g4 / total) * 100.0;

            OnPropertyChanged(nameof(RecordingButtonText));
        }

        private double GroupSum((int scale, int ch)[] g)
        {
            double sum = 0;
            foreach (var (scale, ch) in g)
            {
                int idx = Math.Clamp(ch - 1, 0, 3);
                if (scale == 1) sum += _s1ChKg[idx];
                else if (scale == 2) sum += _s2ChKg[idx];
            }
            return sum;
        }

        private string FormatKg(double kg) => kg.ToString("0.00") + " kg";

        private void UpdateCanConnect()
        {
            OnPropertyChanged(nameof(CanConnectScale1));
            OnPropertyChanged(nameof(CanConnectScale2));
            OnPropertyChanged(nameof(CanConnectAny));
            OnPropertyChanged(nameof(CanDisconnectAny));
            OnPropertyChanged(nameof(CanZeroBoth));

            ConnectScale1Command.RaiseCanExecuteChanged();
            DisconnectScale1Command.RaiseCanExecuteChanged();
            ConnectScale2Command.RaiseCanExecuteChanged();
            DisconnectScale2Command.RaiseCanExecuteChanged();

            ConnectBothCommand.RaiseCanExecuteChanged();
            DisconnectBothCommand.RaiseCanExecuteChanged();

            ZeroScale1Command.RaiseCanExecuteChanged();
            ZeroScale2Command.RaiseCanExecuteChanged();
            ZeroBothCommand.RaiseCanExecuteChanged();
        }

        private void OnPropertyChanged([CallerMemberName] string? name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
