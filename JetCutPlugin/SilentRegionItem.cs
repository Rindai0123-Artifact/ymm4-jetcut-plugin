using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace JetCutPlugin;

/// <summary>
/// UI上で選択/解除可能な無音区間
/// </summary>
public class SilentRegionItem : INotifyPropertyChanged
{
    public event PropertyChangedEventHandler? PropertyChanged;

    public int Index { get; set; }
    public string SourceFile { get; set; } = "";
    public TimeSpan Start { get; set; }
    public TimeSpan End { get; set; }
    public TimeSpan Duration => End - Start;

    private bool _isSelected = true;
    public bool IsSelected
    {
        get => _isSelected;
        set { _isSelected = value; OnPropertyChanged(); }
    }

    public string DisplayText =>
        $"[{Index,3}]  {Fmt(Start)} ～ {Fmt(End)}   ({Duration.TotalSeconds:F1}秒)   {System.IO.Path.GetFileName(SourceFile)}";

    public AudioAnalyzer.SilentRegion ToSilentRegion() => new(Start, End);

    private static string Fmt(TimeSpan t)
    {
        if (t.TotalHours >= 1) return $"{(int)t.TotalHours}:{t.Minutes:D2}:{t.Seconds:D2}";
        return $"{t.Minutes}:{t.Seconds:D2}.{t.Milliseconds / 100}";
    }

    protected void OnPropertyChanged([CallerMemberName] string? n = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(n));
}
