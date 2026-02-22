using System.Diagnostics;
using System.IO;
using System.Reflection;

namespace JetCutPlugin;

/// <summary>
/// YMM4で現在開いているプロジェクトの情報を取得するヘルパー
/// Application.Current.MainWindow.DataContext（= MainModel）経由でアクセス
/// </summary>
public static class ProjectDetector
{
    private static object? _cachedMainModel;
    private static Type? _cachedMainModelType;

    /// <summary>
    /// 現在開いているプロジェクトのファイルパスを取得
    /// </summary>
    public static string? GetCurrentProjectPath()
    {
        // 方法1: MainModel.ProjectFilePath をリフレクションで取得
        var path = GetPropertyValue<string>("ProjectFilePath");
        if (!string.IsNullOrEmpty(path) && File.Exists(path))
            return path;

        // 方法2: MainWindowのタイトルから推測
        path = GetProjectPathFromTitle();
        if (!string.IsNullOrEmpty(path) && File.Exists(path))
            return path;

        // 方法3: user\project フォルダから最新のymmpを探す
        path = FindLatestYmmpFile();
        if (!string.IsNullOrEmpty(path))
            return path;

        return null;
    }

    /// <summary>
    /// 現在の再生フレーム位置を取得
    /// </summary>
    public static int GetCurrentFrame()
    {
        // CurrentFrame は ReactiveProperty<int> の可能性あり
        var raw = GetPropertyValueRaw("CurrentFrame");
        if (raw == null) return 0;

        if (raw is int frame) return frame;

        // ReactiveProperty<T>.Value 経由
        try
        {
            var valueProp = raw.GetType().GetProperty("Value");
            if (valueProp != null)
            {
                var val = valueProp.GetValue(raw);
                if (val is int fv) return fv;
                if (val is long lv) return (int)lv;
            }
        }
        catch { }

        return 0;
    }

    /// <summary>
    /// MainModelのプロパティ値を取得（型付き）
    /// </summary>
    private static T? GetPropertyValue<T>(string propertyName) where T : class
    {
        var raw = GetPropertyValueRaw(propertyName);
        if (raw is T typed) return typed;

        // ReactiveProperty<T>.Value 経由
        try
        {
            var valueProp = raw?.GetType().GetProperty("Value");
            if (valueProp != null)
                return valueProp.GetValue(raw) as T;
        }
        catch { }

        return default;
    }

    /// <summary>
    /// MainModelのプロパティをリフレクションで取得
    /// </summary>
    private static object? GetPropertyValueRaw(string propertyName)
    {
        try
        {
            var (model, modelType) = GetMainModel();
            if (model == null || modelType == null) return null;

            var prop = modelType.GetProperty(propertyName,
                BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
            return prop?.GetValue(model);
        }
        catch { return null; }
    }

    /// <summary>
    /// Application.Current.MainWindow.DataContext からMainModelを取得
    /// </summary>
    private static (object? Model, Type? ModelType) GetMainModel()
    {
        if (_cachedMainModel != null && _cachedMainModelType != null)
            return (_cachedMainModel, _cachedMainModelType);

        try
        {
            // WPF Application.Current から MainWindow の DataContext を取得
            var appType = Type.GetType("System.Windows.Application, PresentationFramework");
            var currentProp = appType?.GetProperty("Current", BindingFlags.Public | BindingFlags.Static);
            var app = currentProp?.GetValue(null);
            if (app == null) return (null, null);

            var mainWindowProp = appType!.GetProperty("MainWindow", BindingFlags.Public | BindingFlags.Instance);
            var mainWindow = mainWindowProp?.GetValue(app);
            if (mainWindow == null) return (null, null);

            var dcProp = mainWindow.GetType().GetProperty("DataContext",
                BindingFlags.Public | BindingFlags.Instance);
            var dataContext = dcProp?.GetValue(mainWindow);
            if (dataContext == null) return (null, null);

            _cachedMainModel = dataContext;
            _cachedMainModelType = dataContext.GetType();
            return (_cachedMainModel, _cachedMainModelType);
        }
        catch { return (null, null); }
    }

    /// <summary>
    /// ウィンドウタイトルからプロジェクトパスを推測
    /// </summary>
    private static string? GetProjectPathFromTitle()
    {
        try
        {
            var appType = Type.GetType("System.Windows.Application, PresentationFramework");
            var app = appType?.GetProperty("Current", BindingFlags.Public | BindingFlags.Static)?.GetValue(null);
            var mainWindow = appType?.GetProperty("MainWindow")?.GetValue(app);
            var title = mainWindow?.GetType().GetProperty("Title")?.GetValue(mainWindow) as string;
            if (string.IsNullOrEmpty(title)) return null;

            // "プロジェクト名 - ゆっくりMovieMaker4" のパターン
            var dash = title.LastIndexOf(" - ");
            if (dash <= 0) return null;

            var name = title[..dash].Trim();

            // フルパスの場合
            if (name.EndsWith(".ymmp", StringComparison.OrdinalIgnoreCase) && File.Exists(name))
                return name;

            // YMM4ディレクトリのuser\projectから探す
            var ymm4Dir = GetYmm4Dir();
            if (ymm4Dir == null) return null;

            var projectDir = Path.Combine(ymm4Dir, "user", "project");
            if (!Directory.Exists(projectDir)) return null;

            foreach (var f in Directory.GetFiles(projectDir, "*.ymmp", SearchOption.AllDirectories))
            {
                if (Path.GetFileNameWithoutExtension(f).Equals(name, StringComparison.OrdinalIgnoreCase))
                    return f;
            }
        }
        catch { }
        return null;
    }

    /// <summary>
    /// 最新のymmpファイルを探す
    /// </summary>
    private static string? FindLatestYmmpFile()
    {
        try
        {
            var ymm4Dir = GetYmm4Dir();
            if (ymm4Dir == null) return null;

            var projectDir = Path.Combine(ymm4Dir, "user", "project");
            if (!Directory.Exists(projectDir)) return null;

            return Directory.GetFiles(projectDir, "*.ymmp", SearchOption.AllDirectories)
                .Select(f => new FileInfo(f))
                .OrderByDescending(f => f.LastWriteTime)
                .FirstOrDefault()?.FullName;
        }
        catch { return null; }
    }

    private static string? GetYmm4Dir()
    {
        try { return Path.GetDirectoryName(Process.GetCurrentProcess().MainModule?.FileName); }
        catch { return null; }
    }
}
