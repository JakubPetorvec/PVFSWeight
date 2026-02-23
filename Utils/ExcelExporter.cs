using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;

using VUKVWeightApp.Domain;
using VUKVWeightApp.ViewModels;

namespace VUKVWeightApp.Utils
{
    public static class ExcelExporter
    {
        // -----------------------------
        // Výpočty
        // -----------------------------
        private const double ActiveThresholdKg = 5.0;
        private const int MedianWindow = 7;
        private const double PlateauRangeFactor = 0.06;
        private const double TakeMiddleFraction = 0.60;
        private const int MinStableSamples = 12;
        private const double MadK = 3.5;

        // -----------------------------
        // Layout
        // -----------------------------
        private const int MarginCols = 1;
        private const int MarginRows = 1;

        private const double DashColWidth = 10.0;

        // grafy: nech čitelné, ale užší dashboard
        private const int ChartW = 460;
        private const int ChartH = 255;

        // popisky osy X (jen každý N-tý)
        private const int XLabelEvery = 4;

        // mezera mezi horním a spodním grafem (v řádcích/buňkách)
        private const int ChartsVerticalGapRows = 1;

        // jak blízko mají být levý a pravý sloupec grafů
        // menší = blíž k sobě
        private const int RightChartsColOffset = 8; // bylo 9 -> posun doprava menší

        public static void ExportSamples(string filePath, IEnumerable<SampleRow> samples)
        {
            ExcelPackage.License.SetNonCommercialOrganization("WeightApp");

            var list = samples?.ToList() ?? new List<SampleRow>();
            if (list.Count == 0)
                throw new InvalidOperationException("Nejsou k dispozici žádná data pro export.");

            using var package = new ExcelPackage();

            var wsData = package.Workbook.Worksheets.Add("Data");
            BuildDataSheet(wsData, list);

            var wsDash = package.Workbook.Worksheets.Add("Dashboard");
            BuildDashboardSheet(wsDash, wsData, list);

            package.Workbook.Worksheets.MoveToStart("Dashboard");
            package.SaveAs(new FileInfo(filePath));
        }

        // ============================================================
        // DATA
        // ============================================================
        private static void BuildDataSheet(ExcelWorksheet ws, List<SampleRow> list)
        {
            var headers = new[]
            {
                "Index",
                "RX Váha 1",
                "RX Váha 2",
                "Celkem [kg]",
                "Váha 1 [kg]",
                "Váha 2 [kg]",
                "G1 [kg]",
                "G2 [kg]",
                "G3 [kg]",
                "G4 [kg]",
                "X label"
            };

            for (int c = 0; c < headers.Length; c++)
                ws.Cells[1, c + 1].Value = headers[c];

            using (var r = ws.Cells[1, 1, 1, headers.Length])
            {
                r.Style.Font.Bold = true;
                r.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                r.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            }

            int row = 2;
            int idx = 1;

            foreach (var s in list)
            {
                ws.Cells[row, 1].Value = idx;

                ws.Cells[row, 2].Value = s.Scale1RxTime;
                ws.Cells[row, 3].Value = s.Scale2RxTime;

                ws.Cells[row, 4].Value = ToWholeKg_NoScale(s.TotalKg);
                ws.Cells[row, 5].Value = ToWholeKg_NoScale(s.Scale1Kg);
                ws.Cells[row, 6].Value = ToWholeKg_NoScale(s.Scale2Kg);

                ws.Cells[row, 7].Value = ToWholeKg_NoScale(s.G1Kg);
                ws.Cells[row, 8].Value = ToWholeKg_NoScale(s.G2Kg);
                ws.Cells[row, 9].Value = ToWholeKg_NoScale(s.G3Kg);
                ws.Cells[row, 10].Value = ToWholeKg_NoScale(s.G4Kg);

                ws.Cells[row, 11].Value = (idx % XLabelEvery == 0) ? idx.ToString(CultureInfo.InvariantCulture) : "";

                idx++;
                row++;
            }

            using (var r = ws.Cells[2, 4, row - 1, 10])
            {
                r.Style.Numberformat.Format = "0";
                r.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            }

            ws.View.FreezePanes(2, 1);
            ws.Cells[ws.Dimension.Address].AutoFitColumns();
        }

        // ============================================================
        // DASHBOARD
        // ============================================================
        private static void BuildDashboardSheet(ExcelWorksheet dash, ExcelWorksheet data, List<SampleRow> list)
        {
            dash.View.ShowGridLines = false;

            int r0 = 1 + MarginRows;
            int c0 = 1 + MarginCols;

            for (int c = 1; c <= 18; c++)
                dash.Column(c).Width = DashColWidth;

            dash.Row(1).Height = 8;

            // ---- Header (méně roztáhlé doprava) ----
            dash.Cells[r0, c0].Value = "PVFS – Dashboard export";
            dash.Cells[r0, c0, r0, c0 + 9].Merge = true; // bylo +10
            dash.Cells[r0, c0].Style.Font.Bold = true;
            dash.Cells[r0, c0].Style.Font.Size = 16;

            dash.Cells[r0 + 1, c0].Value = $"Vygenerováno: {DateTime.Now:dd.MM.yyyy HH:mm:ss}";
            dash.Cells[r0 + 1, c0, r0 + 1, c0 + 9].Merge = true; // bylo +10
            dash.Cells[r0 + 1, c0].Style.Font.Size = 10;

            var window = ComputeStableWindow(list, s => ParseKg_NoScale(s.TotalKg));
            dash.Cells[r0 + 2, c0].Value = $"Okno průměru: index {window.Start + 1}–{window.End + 1} (N={window.Length})";
            dash.Cells[r0 + 2, c0, r0 + 2, c0 + 9].Merge = true; // bylo +10
            dash.Cells[r0 + 2, c0].Style.Font.Size = 10;
            dash.Cells[r0 + 2, c0].Style.Font.Color.SetColor(System.Drawing.Color.DimGray);

            // KPI
            var avgTotal = AverageInWindow(list, s => ParseKg_NoScale(s.TotalKg), window);
            var avgS1 = AverageInWindow(list, s => ParseKg_NoScale(s.Scale1Kg), window);
            var avgS2 = AverageInWindow(list, s => ParseKg_NoScale(s.Scale2Kg), window);

            var avgG1 = AverageInWindow(list, s => ParseKg_NoScale(s.G1Kg), window);
            var avgG2 = AverageInWindow(list, s => ParseKg_NoScale(s.G2Kg), window);
            var avgG3 = AverageInWindow(list, s => ParseKg_NoScale(s.G3Kg), window);
            var avgG4 = AverageInWindow(list, s => ParseKg_NoScale(s.G4Kg), window);

            var lastActive = FindLastActiveSample(list, s => ParseKg_NoScale(s.TotalKg), ActiveThresholdKg) ?? list[^1];

            // horní řada KPI (pravý okraj stáhnu o 1 sloupec doleva)
            DrawKpiCard(dash, Addr(r0 + 4, c0 + 0, r0 + 5, c0 + 3), "Celkem (avg)", RoundKg(avgTotal), "kg", extraLine: $"N={window.Length}");
            DrawKpiCard(dash, Addr(r0 + 4, c0 + 4, r0 + 5, c0 + 7), "Váha 1 (avg)", RoundKg(avgS1), "kg");
            DrawKpiCard(dash, Addr(r0 + 4, c0 + 8, r0 + 5, c0 + 11), "Váha 2 (avg)", RoundKg(avgS2), "kg");
            DrawRxCard(dash, Addr(r0 + 4, c0 + 12, r0 + 5, c0 + 15), $"{lastActive.Scale1RxTime} | {lastActive.Scale2RxTime}"); // bylo do +16

            // spodní řada KPI
            DrawKpiCard(dash, Addr(r0 + 6, c0 + 0, r0 + 7, c0 + 3), "G1 (avg)", RoundKg(avgG1), "kg");
            DrawKpiCard(dash, Addr(r0 + 6, c0 + 4, r0 + 7, c0 + 7), "G2 (avg)", RoundKg(avgG2), "kg");
            DrawKpiCard(dash, Addr(r0 + 6, c0 + 8, r0 + 7, c0 + 11), "G3 (avg)", RoundKg(avgG3), "kg");
            DrawKpiCard(dash, Addr(r0 + 6, c0 + 12, r0 + 7, c0 + 15), "G4 (avg)", RoundKg(avgG4), "kg"); // bylo do +16

            // Charts (X labels z "X label")
            int xFrom = 2;
            int xTo = list.Count + 1;
            int xLabelCol = 11;

            int chartsRowTop = r0 + 9;

            int leftChartCol = c0 + 0;
            int rightChartCol = c0 + RightChartsColOffset; // BLÍŽ k sobě

            // horní řada grafů
            AddChart(dash, eChartType.Line, "Celková váha (Celkem)", data, xFrom, xTo, xLabelCol,
                new[] { 4 }, new[] { "Celkem [kg]" },
                fromRow: chartsRowTop, fromCol: leftChartCol, width: ChartW, height: ChartH);

            AddChart(dash, eChartType.Line, "Váha 1 a Váha 2", data, xFrom, xTo, xLabelCol,
                new[] { 5, 6 }, new[] { "Váha 1 [kg]", "Váha 2 [kg]" },
                fromRow: chartsRowTop, fromCol: rightChartCol, width: ChartW, height: ChartH);

            // spodní řada grafů – posun o ChartH "na buňky" neumíme přímo,
            // ale v řádcích dáme explicitní mezeru 1 řádek (ChartsVerticalGapRows)
            int chartsRowBottom = chartsRowTop + 12 + ChartsVerticalGapRows; // bylo +12 natěsno

            AddChart(dash, eChartType.Line, "Skupiny (G1–G4)", data, xFrom, xTo, xLabelCol,
                new[] { 7, 8, 9, 10 }, new[] { "G1 [kg]", "G2 [kg]", "G3 [kg]", "G4 [kg]" },
                fromRow: chartsRowBottom, fromCol: leftChartCol, width: ChartW, height: ChartH);

            AddChart(dash, eChartType.AreaStacked, "Podíl skupin (G1–G4)", data, xFrom, xTo, xLabelCol,
                new[] { 7, 8, 9, 10 }, new[] { "G1", "G2", "G3", "G4" },
                fromRow: chartsRowBottom, fromCol: rightChartCol, width: ChartW, height: ChartH);

            dash.View.ZoomScale = 100;
        }

        private static string Addr(int r1, int c1, int r2, int c2) =>
            new ExcelAddress(r1, c1, r2, c2).Address;

        // ============================================================
        // KPI
        // ============================================================
        private static void DrawKpiCard(ExcelWorksheet ws, string rangeA1, string title, int value, string unit, string? extraLine = null)
        {
            var addr = new ExcelAddress(rangeA1);
            int top = addr.Start.Row;
            int left = addr.Start.Column;
            int bottom = addr.End.Row;
            int right = addr.End.Column;

            using (var card = ws.Cells[top, left, bottom, right])
            {
                card.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                card.Style.Font.Name = "Segoe UI";
                card.Style.Fill.PatternType = ExcelFillStyle.Solid;
                card.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(245, 248, 255));
                card.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            }

            ws.Cells[top, left].Value = title;
            ws.Cells[top, left].Style.Font.Bold = true;
            ws.Cells[top, left].Style.Font.Size = 11;

            ws.Cells[top + 1, left].Value = $"{value} {unit}".Trim();
            ws.Cells[top + 1, left].Style.Font.Bold = true;
            ws.Cells[top + 1, left].Style.Font.Size = 16;

            if (!string.IsNullOrWhiteSpace(extraLine))
            {
                ws.Cells[top + 1, left + 2].Value = extraLine;
                ws.Cells[top + 1, left + 2].Style.Font.Size = 10;
                ws.Cells[top + 1, left + 2].Style.Font.Color.SetColor(System.Drawing.Color.DimGray);
            }
        }

        private static void DrawRxCard(ExcelWorksheet ws, string rangeA1, string extraLine)
        {
            var addr = new ExcelAddress(rangeA1);
            int top = addr.Start.Row;
            int left = addr.Start.Column;
            int bottom = addr.End.Row;
            int right = addr.End.Column;

            using (var card = ws.Cells[top, left, bottom, right])
            {
                card.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                card.Style.Font.Name = "Segoe UI";
                card.Style.Fill.PatternType = ExcelFillStyle.Solid;
                card.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(245, 248, 255));
                card.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            }

            ws.Cells[top, left].Value = "RX";
            ws.Cells[top, left].Style.Font.Bold = true;
            ws.Cells[top, left].Style.Font.Size = 11;

            ws.Cells[top + 1, left].Value = extraLine;
            ws.Cells[top + 1, left].Style.Font.Size = 10;
            ws.Cells[top + 1, left].Style.Font.Color.SetColor(System.Drawing.Color.DimGray);
        }

        // ============================================================
        // CHARTS
        // ============================================================
        private static void AddChart(
            ExcelWorksheet dash,
            eChartType chartType,
            string title,
            ExcelWorksheet dataWs,
            int xFromRow,
            int xToRow,
            int xLabelCol,
            int[] yCols,
            string[] seriesNames,
            int fromRow,
            int fromCol,
            int width,
            int height)
        {
            var chart = dash.Drawings.AddChart(Guid.NewGuid().ToString("N"), chartType);
            if (chart == null) return;

            chart.Title.Text = title;
            chart.Legend.Position = eLegendPosition.Bottom;

            var xRange = dataWs.Cells[xFromRow, xLabelCol, xToRow, xLabelCol];

            for (int i = 0; i < yCols.Length; i++)
            {
                var yRange = dataWs.Cells[xFromRow, yCols[i], xToRow, yCols[i]];
                var s = chart.Series.Add(yRange, xRange);
                s.Header = seriesNames[i];
            }

            // Y osa: kg
            try { chart.YAxis.Title.Text = "kg"; } catch { }

            chart.SetPosition(fromRow - 1, 0, fromCol - 1, 0);
            chart.SetSize(width, height);
        }

        // ============================================================
        // STABLE WINDOW
        // ============================================================
        private readonly struct Window
        {
            public Window(int start, int end) { Start = start; End = end; }
            public int Start { get; }
            public int End { get; }
            public int Length => Math.Max(0, End - Start + 1);
        }

        private static Window ComputeStableWindow(List<SampleRow> list, Func<SampleRow, double> selectorKg)
        {
            var raw = list.Select(selectorKg).ToArray();
            if (raw.Length == 0) return new Window(0, 0);

            var seg = FindLastActiveSegment(raw, ActiveThresholdKg);
            int segStart = seg.start;
            int segEnd = seg.end;

            var segmentRaw = raw.Skip(segStart).Take(segEnd - segStart + 1).ToArray();
            if (segmentRaw.Length == 0) return new Window(0, raw.Length - 1);

            var smooth = MedianFilter(segmentRaw, MedianWindow);

            int maxIdx = ArgMax(smooth);
            double max = smooth[maxIdx];
            double min = smooth.Min();
            double range = Math.Max(1e-9, max - min);

            double threshold = max - (range * PlateauRangeFactor);

            int left = maxIdx;
            while (left > 0 && smooth[left - 1] >= threshold) left--;

            int right = maxIdx;
            while (right < smooth.Length - 1 && smooth[right + 1] >= threshold) right++;

            int plateauLen = right - left + 1;

            if (plateauLen < MinStableSamples)
            {
                double threshold2 = max - (range * (PlateauRangeFactor * 1.8));
                left = maxIdx;
                while (left > 0 && smooth[left - 1] >= threshold2) left--;
                right = maxIdx;
                while (right < smooth.Length - 1 && smooth[right + 1] >= threshold2) right++;
                plateauLen = right - left + 1;
            }

            int wStartSeg, wEndSeg;

            if (plateauLen >= MinStableSamples)
            {
                int takeLen = (int)Math.Round(plateauLen * TakeMiddleFraction);
                takeLen = Math.Max(MinStableSamples, Math.Min(plateauLen, takeLen));

                int mid = left + plateauLen / 2;
                int half = takeLen / 2;

                wStartSeg = Math.Max(left, mid - half);
                wEndSeg = wStartSeg + takeLen - 1;
                if (wEndSeg > right)
                {
                    wEndSeg = right;
                    wStartSeg = Math.Max(left, wEndSeg - takeLen + 1);
                }
            }
            else
            {
                int n = segmentRaw.Length;
                wStartSeg = n / 3;
                wEndSeg = (2 * n / 3) - 1;
                if (wEndSeg <= wStartSeg) { wStartSeg = 0; wEndSeg = n - 1; }
            }

            int absStart = segStart + wStartSeg;
            int absEnd = segStart + wEndSeg;

            absStart = Math.Max(0, Math.Min(absStart, raw.Length - 1));
            absEnd = Math.Max(0, Math.Min(absEnd, raw.Length - 1));
            if (absEnd < absStart) (absStart, absEnd) = (absEnd, absStart);

            return new Window(absStart, absEnd);
        }

        private static double AverageInWindow(List<SampleRow> list, Func<SampleRow, double> selectorKg, Window w)
        {
            if (w.Length <= 0) return 0;

            var values = list.Skip(w.Start).Take(w.Length).Select(selectorKg).ToArray();
            if (values.Length == 0) return 0;

            var filtered = MadFilter(values, MadK);
            var used = filtered.Length >= Math.Max(5, values.Length / 2) ? filtered : values;

            return used.Average();
        }

        // ============================================================
        // ACTIVE SEGMENT
        // ============================================================
        private static SampleRow? FindLastActiveSample(List<SampleRow> list, Func<SampleRow, double> selectorKg, double thresholdKg)
        {
            for (int i = list.Count - 1; i >= 0; i--)
            {
                if (selectorKg(list[i]) >= thresholdKg)
                    return list[i];
            }
            return null;
        }

        private static (int start, int end) FindLastActiveSegment(double[] rawKg, double thresholdKg)
        {
            int end = -1;
            for (int i = rawKg.Length - 1; i >= 0; i--)
            {
                if (rawKg[i] >= thresholdKg) { end = i; break; }
            }

            if (end < 0)
                return (0, rawKg.Length - 1);

            int start = end;
            int holes = 0;
            const int maxHoles = 3;

            while (start > 0)
            {
                if (rawKg[start - 1] >= thresholdKg)
                {
                    holes = 0;
                    start--;
                }
                else
                {
                    holes++;
                    if (holes > maxHoles) break;
                    start--;
                }
            }

            while (start < end && rawKg[start] < thresholdKg) start++;
            return (start, end);
        }

        // ============================================================
        // Filters
        // ============================================================
        private static double[] MedianFilter(double[] data, int window)
        {
            if (data.Length == 0) return Array.Empty<double>();
            if (window <= 1) return (double[])data.Clone();
            if (window % 2 == 0) window++;

            int half = window / 2;
            var result = new double[data.Length];

            for (int i = 0; i < data.Length; i++)
            {
                int a = Math.Max(0, i - half);
                int b = Math.Min(data.Length - 1, i + half);
                int len = b - a + 1;

                var tmp = new double[len];
                Array.Copy(data, a, tmp, 0, len);
                Array.Sort(tmp);
                result[i] = tmp[len / 2];
            }

            return result;
        }

        private static int ArgMax(double[] data)
        {
            int idx = 0;
            double best = double.NegativeInfinity;
            for (int i = 0; i < data.Length; i++)
            {
                if (data[i] > best)
                {
                    best = data[i];
                    idx = i;
                }
            }
            return idx;
        }

        private static double[] MadFilter(double[] values, double k)
        {
            if (values.Length == 0) return Array.Empty<double>();
            if (values.Length < 6) return values;

            double median = Median(values);
            var absDev = values.Select(v => Math.Abs(v - median)).ToArray();
            double mad = Median(absDev);

            if (mad < 1e-12) return values;

            double sigma = 1.4826 * mad;
            double lo = median - (k * sigma);
            double hi = median + (k * sigma);

            return values.Where(v => v >= lo && v <= hi).ToArray();
        }

        private static double Median(double[] values)
        {
            var tmp = (double[])values.Clone();
            Array.Sort(tmp);
            int n = tmp.Length;
            if (n % 2 == 1) return tmp[n / 2];
            return 0.5 * (tmp[(n / 2) - 1] + tmp[n / 2]);
        }

        // ============================================================
        // Parsing / units (NO SCALE)
        // ============================================================
        private static int RoundKg(double v) =>
            (int)Math.Round(v, MidpointRounding.AwayFromZero);

        private static int ToWholeKg_NoScale(string? text) =>
            RoundKg(ParseKg_NoScale(text));

        private static double ParseKg_NoScale(string? text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return 0;

            var t = text.Replace("kg", "", StringComparison.OrdinalIgnoreCase)
                        .Replace(" ", "")
                        .Trim();

            if (double.TryParse(t, NumberStyles.Any, CultureInfo.CurrentCulture, out var v))
                return v;

            if (double.TryParse(t, NumberStyles.Any, CultureInfo.InvariantCulture, out v))
                return v;

            t = t.Replace(",", ".");
            if (double.TryParse(t, NumberStyles.Any, CultureInfo.InvariantCulture, out v))
                return v;

            return 0;
        }
    }
}
