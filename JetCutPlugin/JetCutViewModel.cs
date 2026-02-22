using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using Microsoft.Win32;

namespace JetCutPlugin;

/// <summary>
/// ジェットカットツールのViewModel v4
/// ・現在のプロジェクトを自動検出 + 複数プロジェクト手動追加
/// ・再生位置（秒）以降のメディアを解析
/// ・無音区間の選択的カット
/// </summary>
public class JetCutViewModel : INotifyPropertyChanged
{
    public event PropertyChangedEventHandler? PropertyChanged;

    // ==== プロジェクトファイル一覧 ====

    public ObservableCollection<ProjectFileItem> Projects { get; } = new();

    private ProjectFileItem? _selectedProject;
    public ProjectFileItem? SelectedProject
    {
        get => _selectedProject;
        set { _selectedProject = value; OnPropertyChanged(); }
    }

    public string ProjectSummary => Projects.Count switch
    {
        0 => "プロジェクトがありません",
        1 => $"📂 {Path.GetFileName(Projects[0].FilePath)}",
        _ => $"📂 {Projects.Count}件のプロジェクト"
    };

    // ==== メディアファイル ====
    public ObservableCollection<MediaFileItem> MediaFiles { get; } = new();

    // ==== 再生位置 ====

    private double _startPositionSec;
    public double StartPositionSec
    {
        get => _startPositionSec;
        set { _startPositionSec = Math.Max(0, value); OnPropertyChanged(); OnPropertyChanged(nameof(StartPositionText)); }
    }
    public string StartPositionText => FormatTimeHMS(TimeSpan.FromSeconds(StartPositionSec));

    // ==== パラメータ ====

    private double _thresholdDb = -40.0;
    public double ThresholdDb
    {
        get => _thresholdDb;
        set { _thresholdDb = value; OnPropertyChanged(); OnPropertyChanged(nameof(ThresholdDbText)); }
    }
    public string ThresholdDbText => $"{ThresholdDb:F0} dB";

    private double _minSilenceSec = 0.5;
    public double MinSilenceSec
    {
        get => _minSilenceSec;
        set { _minSilenceSec = Math.Round(value, 1); OnPropertyChanged(); OnPropertyChanged(nameof(MinSilenceText)); }
    }
    public string MinSilenceText => $"{MinSilenceSec:F1} 秒";

    private double _marginSec = 0.1;
    public double MarginSec
    {
        get => _marginSec;
        set { _marginSec = Math.Round(value, 2); OnPropertyChanged(); OnPropertyChanged(nameof(MarginText)); }
    }
    public string MarginText => $"{MarginSec:F2} 秒";

    // ==== 無音区間 ====
    public ObservableCollection<SilentRegionItem> SilentRegions { get; } = new();

    private int SelectedCount => SilentRegions.Count(r => r.IsSelected);
    public string SelectionSummary
    {
        get
        {
            if (SilentRegions.Count == 0) return "";
            var selDur = TimeSpan.FromTicks(SilentRegions.Where(r => r.IsSelected).Sum(r => r.Duration.Ticks));
            return $"{SelectedCount} / {SilentRegions.Count} 件選択  （合計 {FormatTimeHMS(selDur)} 短縮）";
        }
    }

    // ==== 状態 ====

    private double _progress;
    public double Progress { get => _progress; set { _progress = value; OnPropertyChanged(); } }

    private string _statusText = "";
    public string StatusText { get => _statusText; set { _statusText = value; OnPropertyChanged(); } }

    private string _logText = "";
    public string LogText { get => _logText; set { _logText = value; OnPropertyChanged(); } }

    private bool _isProcessing;
    public bool IsProcessing
    {
        get => _isProcessing;
        set { _isProcessing = value; OnPropertyChanged(); OnPropertyChanged(nameof(CanAnalyze)); OnPropertyChanged(nameof(CanCut)); }
    }

    private bool _hasResult;
    public bool HasResult
    {
        get => _hasResult;
        set { _hasResult = value; OnPropertyChanged(); OnPropertyChanged(nameof(CanCut)); }
    }

    private string _resultText = "";
    public string ResultText { get => _resultText; set { _resultText = value; OnPropertyChanged(); } }

    public bool CanAnalyze => !IsProcessing && Projects.Count > 0;
    public bool CanCut => !IsProcessing && HasResult && SelectedCount > 0;

    // ==== 内部 ====
    private readonly Dictionary<string, double> _fpsMap = new();
    private CancellationTokenSource? _cts;

    // ==== コマンド ====
    public ICommand DetectCommand => new RelayCommand(_ => DetectCurrentProject());
    public ICommand AddCommand => new RelayCommand(_ => AddProjects());
    public ICommand RemoveCommand => new RelayCommand(_ => RemoveProject(), _ => SelectedProject != null);
    public ICommand AnalyzeCommand => new RelayCommand(async _ => await AnalyzeAsync(), _ => CanAnalyze);
    public ICommand CutCommand => new RelayCommand(async _ => await CutAsync(), _ => CanCut);
    public ICommand CancelCommand => new RelayCommand(_ => _cts?.Cancel(), _ => IsProcessing);
    public ICommand SelectAllCommand => new RelayCommand(_ => SetAll(true));
    public ICommand DeselectAllCommand => new RelayCommand(_ => SetAll(false));

    // ==== 初期化 ====
    public JetCutViewModel()
    {
        DetectCurrentProject();
    }

    // ==== プロジェクト検出 ====

    private void DetectCurrentProject()
    {
        Log("🔍 プロジェクトを検出中...");

        var path = ProjectDetector.GetCurrentProjectPath();
        if (string.IsNullOrEmpty(path))
        {
            StatusText = "⚠ プロジェクトが見つかりません。YMM4でプロジェクトを開くか「追加」で選択してください";
            Log("⚠ 自動検出失敗 — 手動で追加してください");
            return;
        }

        // 既に追加済みでなければ追加
        if (Projects.All(p => p.FilePath != path))
        {
            Projects.Add(new ProjectFileItem(path));
            Log($"✅ 検出: {Path.GetFileName(path)}");
        }

        // 再生位置を取得
        var frame = ProjectDetector.GetCurrentFrame();
        if (frame > 0)
        {
            var fps = YmmpEditor.GetFps(path);
            StartPositionSec = Math.Round(frame / fps, 1);
            Log($"📍 再生位置: {FormatTimeHMS(TimeSpan.FromSeconds(StartPositionSec))} (フレーム {frame})");
        }

        RefreshMediaFiles();
        OnPropertyChanged(nameof(ProjectSummary));
        OnPropertyChanged(nameof(CanAnalyze));
        StatusText = "「🔍 解析」ボタンで無音区間を検出します";
    }

    private void AddProjects()
    {
        var dlg = new OpenFileDialog
        {
            Title = "YMM4プロジェクトファイルを選択（複数可）",
            Filter = "YMM4プロジェクト (*.ymmp)|*.ymmp",
            Multiselect = true
        };
        if (dlg.ShowDialog() != true) return;

        foreach (var f in dlg.FileNames)
        {
            if (Projects.All(p => p.FilePath != f))
            {
                Projects.Add(new ProjectFileItem(f));
                Log($"➕ 追加: {Path.GetFileName(f)}");
            }
        }
        RefreshMediaFiles();
        OnPropertyChanged(nameof(ProjectSummary));
        OnPropertyChanged(nameof(CanAnalyze));
        HasResult = false;
        SilentRegions.Clear();
    }

    private void RemoveProject()
    {
        if (SelectedProject == null) return;
        Log($"➖ 削除: {Path.GetFileName(SelectedProject.FilePath)}");
        Projects.Remove(SelectedProject);
        RefreshMediaFiles();
        OnPropertyChanged(nameof(ProjectSummary));
        OnPropertyChanged(nameof(CanAnalyze));
        HasResult = false;
        SilentRegions.Clear();
    }

    private void RefreshMediaFiles()
    {
        MediaFiles.Clear();
        _fpsMap.Clear();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var proj in Projects)
        {
            if (!File.Exists(proj.FilePath)) continue;
            _fpsMap[proj.FilePath] = YmmpEditor.GetFps(proj.FilePath);

            foreach (var mf in YmmpEditor.GetMediaFilePaths(proj.FilePath))
            {
                if (seen.Add(mf))
                    MediaFiles.Add(new MediaFileItem { FilePath = mf, FileName = Path.GetFileName(mf), Status = "—" });
            }
        }
        Log($"🎬 メディアファイル: {MediaFiles.Count}件");
    }

    // ==== 解析 ====

    private async Task AnalyzeAsync()
    {
        IsProcessing = true;
        HasResult = false;
        SilentRegions.Clear();
        Progress = 0;
        _cts = new CancellationTokenSource();

        try
        {
            Log("═══ 解析開始 ═══");
            if (MediaFiles.Count == 0) { StatusText = "⚠ メディアファイルがありません"; return; }

            var startOffset = TimeSpan.FromSeconds(StartPositionSec);
            if (startOffset > TimeSpan.Zero)
                Log($"📍 開始位置: {FormatTimeHMS(startOffset)} 以降を解析");

            var idx = 1;
            for (var i = 0; i < MediaFiles.Count; i++)
            {
                _cts.Token.ThrowIfCancellationRequested();
                var mf = MediaFiles[i];
                mf.Status = "解析中…";
                StatusText = $"解析中: {mf.FileName}  ({i + 1}/{MediaFiles.Count})";

                try
                {
                    var prog = new Progress<double>(p => Progress = (i + p) / MediaFiles.Count * 100);
                    var result = await AudioAnalyzer.AnalyzeAsync(
                        mf.FilePath, ThresholdDb,
                        (int)(MinSilenceSec * 1000), (int)(MarginSec * 1000),
                        prog, _cts.Token);

                    // 開始位置フィルター
                    var regions = startOffset > TimeSpan.Zero
                        ? result.SilentRegions.Where(r => r.End > startOffset).ToList()
                        : result.SilentRegions;

                    mf.Status = $"✅ {regions.Count}箇所";
                    Log($"  {mf.FileName}: {FormatTimeHMS(result.TotalDuration)} / 無音 {regions.Count}箇所");

                    foreach (var r in regions)
                    {
                        var adjustedStart = r.Start < startOffset ? startOffset : r.Start;
                        SilentRegions.Add(new SilentRegionItem
                        {
                            Index = idx++,
                            SourceFile = mf.FilePath,
                            Start = adjustedStart,
                            End = r.End,
                            IsSelected = true
                        });
                    }
                }
                catch (Exception ex) { mf.Status = "❌ 失敗"; Log($"  エラー: {ex.Message}"); }
            }

            if (SilentRegions.Count > 0)
            {
                HasResult = true;
                var total = TimeSpan.FromTicks(SilentRegions.Sum(r => r.Duration.Ticks));
                ResultText = $"🔇 {SilentRegions.Count}箇所の無音区間（合計 {FormatTimeHMS(total)}）\n" +
                             "チェックで選択 → 「✂ カット実行」で反映";

                foreach (var item in SilentRegions)
                    item.PropertyChanged += (_, _) => { OnPropertyChanged(nameof(SelectionSummary)); OnPropertyChanged(nameof(CanCut)); };
                OnPropertyChanged(nameof(SelectionSummary));
            }
            else
            {
                ResultText = "無音区間は見つかりませんでした\nパラメータを変えて再解析してみてください";
            }

            Progress = 100;
            StatusText = HasResult ? $"✅ {SilentRegions.Count}箇所検出。カットする区間を選んでください" : "無音区間なし";
            Log("═══ 解析完了 ═══");
        }
        catch (OperationCanceledException) { StatusText = "キャンセルしました"; Log("⛔ キャンセル"); }
        catch (Exception ex) { StatusText = $"エラー: {ex.Message}"; Log($"❌ {ex.Message}"); }
        finally { IsProcessing = false; _cts?.Dispose(); _cts = null; }
    }

    // ==== カット実行 ====

    private async Task CutAsync()
    {
        var selected = SilentRegions.Where(r => r.IsSelected).Select(r => r.ToSilentRegion()).ToList();
        if (selected.Count == 0) return;

        IsProcessing = true;
        Progress = 0;

        try
        {
            var merged = Merge(selected);
            Log("═══ ジェットカット実行 ═══");
            Log($"選択: {selected.Count}箇所 → マージ後: {merged.Count}箇所");

            var totalResult = new YmmpEditor.CutResult(0, 0, 0, TimeSpan.Zero);
            for (var i = 0; i < Projects.Count; i++)
            {
                var p = Projects[i];
                if (!File.Exists(p.FilePath)) { Log($"⚠ スキップ: {p.DisplayName}"); continue; }

                StatusText = $"カット中: {p.DisplayName}  ({i + 1}/{Projects.Count})";
                Log($"✂ {p.DisplayName}  (バックアップ → .bak)");

                var fps = _fpsMap.GetValueOrDefault(p.FilePath, 30.0);
                var r = await Task.Run(() => YmmpEditor.ExecuteJetCut(p.FilePath, merged, fps, true));

                Log($"  処理: {r.ItemsProcessed} / 分割: {r.ItemsSplit} / 削除: {r.SilentItemsRemoved} / 短縮: {FormatTimeHMS(r.TimeSaved)}");
                totalResult = new(totalResult.ItemsProcessed + r.ItemsProcessed, totalResult.ItemsSplit + r.ItemsSplit,
                    totalResult.SilentItemsRemoved + r.SilentItemsRemoved, totalResult.TimeSaved + r.TimeSaved);
                Progress = (i + 1.0) / Projects.Count * 100;
            }

            Progress = 100;
            StatusText = "✅ 完了！  YMM4でプロジェクトを再読み込みしてください";
            ResultText = $"✅ カット完了\n" +
                         $"処理 {totalResult.ItemsProcessed}件 / 分割 {totalResult.ItemsSplit}件 / 削除 {totalResult.SilentItemsRemoved}件\n" +
                         $"短縮: {FormatTimeHMS(totalResult.TimeSaved)}\n\n" +
                         "⚠ YMM4でプロジェクトを再読み込みしてください";
            HasResult = false;
            Log("═══ 完了 ═══\n※ YMM4でプロジェクトを再読み込みしてください");
        }
        catch (Exception ex)
        {
            StatusText = $"エラー: {ex.Message}";
            Log($"❌ {ex.Message}\n.bakファイルで復元できます");
        }
        finally { IsProcessing = false; }
    }

    // ==== ユーティリティ ====

    private void SetAll(bool sel)
    {
        foreach (var r in SilentRegions) r.IsSelected = sel;
        OnPropertyChanged(nameof(SelectionSummary));
        OnPropertyChanged(nameof(CanCut));
    }

    private void Log(string msg) => LogText += $"[{DateTime.Now:HH:mm:ss}] {msg}\n";

    /// <summary>分:秒 or 時:分:秒 で表示</summary>
    internal static string FormatTimeHMS(TimeSpan t)
    {
        if (t.TotalHours >= 1) return $"{(int)t.TotalHours}時間{t.Minutes:D2}分{t.Seconds:D2}秒";
        if (t.TotalMinutes >= 1) return $"{t.Minutes}分{t.Seconds:D2}秒";
        return $"{t.TotalSeconds:F1}秒";
    }

    private static List<AudioAnalyzer.SilentRegion> Merge(List<AudioAnalyzer.SilentRegion> regions)
    {
        if (regions.Count == 0) return regions;
        var s = regions.OrderBy(r => r.Start).ToList();
        var m = new List<AudioAnalyzer.SilentRegion>();
        var c = s[0];
        for (var i = 1; i < s.Count; i++)
        {
            if (s[i].Start <= c.End) c = new(c.Start, s[i].End > c.End ? s[i].End : c.End);
            else { m.Add(c); c = s[i]; }
        }
        m.Add(c);
        return m;
    }

    protected void OnPropertyChanged([CallerMemberName] string? n = null) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(n));
}

/// <summary>プロジェクトファイル</summary>
public class ProjectFileItem(string filePath)
{
    public string FilePath { get; } = filePath;
    public string DisplayName => Path.GetFileName(FilePath);
    public override string ToString() => DisplayName;
}

/// <summary>メディアファイル</summary>
public class MediaFileItem : INotifyPropertyChanged
{
    public event PropertyChangedEventHandler? PropertyChanged;
    public string FilePath { get; set; } = "";
    public string FileName { get; set; } = "";
    private string _status = "";
    public string Status { get => _status; set { _status = value; PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Status))); } }
}

/// <summary>RelayCommand</summary>
public class RelayCommand(Action<object?> execute, Func<object?, bool>? canExecute = null) : ICommand
{
    public event EventHandler? CanExecuteChanged { add => CommandManager.RequerySuggested += value; remove => CommandManager.RequerySuggested -= value; }
    public bool CanExecute(object? p) => canExecute?.Invoke(p) ?? true;
    public void Execute(object? p) => execute(p);
}
