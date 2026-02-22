using YukkuriMovieMaker.Plugin;

namespace JetCutPlugin;

/// <summary>
/// YMM4 ジェットカット（無音カット）ツールプラグイン
/// タイムライン上の音声/動画アイテムの無音区間を検出し、自動でカットします。
/// </summary>
public class JetCutToolPlugin : IToolPlugin
{
    /// <summary>プラグイン名</summary>
    public string Name => "ジェットカット（無音カット）";

    /// <summary>ViewModelの型</summary>
    public Type ViewModelType => typeof(JetCutViewModel);

    /// <summary>Viewの型</summary>
    public Type ViewType => typeof(JetCutView);
}
