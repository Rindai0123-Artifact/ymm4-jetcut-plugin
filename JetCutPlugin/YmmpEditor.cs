using System.IO;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace JetCutPlugin;

/// <summary>音声/動画のファイル拡張子セット</summary>
static class MediaExtensions
{
    public static readonly HashSet<string> Supported = new(StringComparer.OrdinalIgnoreCase)
    {
        ".mp4", ".avi", ".mov", ".mkv", ".wmv",
        ".mp3", ".wav", ".ogg", ".flac", ".aac", ".m4a"
    };
}

/// <summary>
/// YMM4のymmpプロジェクトファイル(JSON)を読み書きし、
/// 無音区間に基づいてタイムラインアイテムを分割・削除するクラス
/// </summary>
public class YmmpEditor
{
    /// <summary>カット結果</summary>
    public record CutResult(int ItemsProcessed, int ItemsSplit, int SilentItemsRemoved, TimeSpan TimeSaved);

    /// <summary>アイテム情報</summary>
    public record ItemInfo(string Type, int Frame, int Length, int Layer, bool HasAudioFile);

    /// <summary>
    /// ymmpファイルを読み込み、含まれる音声/動画アイテムの情報を取得します
    /// </summary>
    public static List<ItemInfo> GetMediaItems(string ymmpPath)
    {
        var json = File.ReadAllText(ymmpPath);
        var doc = JsonNode.Parse(json);
        var items = new List<ItemInfo>();

        if (doc == null) return items;

        var timelines = doc["Timelines"]?.AsArray();
        if (timelines == null) return items;

        foreach (var timeline in timelines)
        {
            var timelineItems = timeline?["Items"]?.AsArray();
            if (timelineItems == null) continue;

            foreach (var item in timelineItems)
            {
                if (item == null) continue;

                var type = item["$type"]?.GetValue<string>() ?? "";
                var frame = item["Frame"]?.GetValue<int>() ?? 0;
                var length = item["Length"]?.GetValue<int>() ?? 0;
                var layer = item["Layer"]?.GetValue<int>() ?? 0;

                // 音声/動画アイテムかどうか判定
                var hasAudio = HasAudioContent(item, type);
                if (hasAudio)
                {
                    items.Add(new ItemInfo(type, frame, length, layer, true));
                }
            }
        }

        return items;
    }

    /// <summary>
    /// ymmpファイルに対してジェットカットを実行します
    /// </summary>
    /// <param name="ymmpPath">ymmpファイルのパス</param>
    /// <param name="silentRegions">無音区間のリスト</param>
    /// <param name="fps">プロジェクトのFPS</param>
    /// <param name="createBackup">バックアップを作成するか</param>
    /// <returns>カット結果</returns>
    public static CutResult ExecuteJetCut(
        string ymmpPath,
        List<AudioAnalyzer.SilentRegion> silentRegions,
        double fps,
        bool createBackup = true)
    {
        // バックアップ作成
        if (createBackup)
        {
            var backupPath = ymmpPath + ".bak";
            File.Copy(ymmpPath, backupPath, overwrite: true);
        }

        var json = File.ReadAllText(ymmpPath);
        var doc = JsonNode.Parse(json);
        if (doc == null) return new CutResult(0, 0, 0, TimeSpan.Zero);

        var fps_val = fps;
        
        // プロジェクトからFPS情報を取得（指定がなければ）
        if (fps_val <= 0)
        {
            fps_val = GetProjectFps(doc) ?? 30.0;
        }

        var itemsProcessed = 0;
        var itemsSplit = 0;
        var silentItemsRemoved = 0;
        var totalFramesSaved = 0L;

        var timelines = doc["Timelines"]?.AsArray();
        if (timelines == null) return new CutResult(0, 0, 0, TimeSpan.Zero);

        foreach (var timeline in timelines)
        {
            var timelineItems = timeline?["Items"]?.AsArray();
            if (timelineItems == null) continue;

            var newItems = new JsonArray();
            var frameOffset = 0; // 削除による累計フレームオフセット

            // 全アイテムをフレーム順でソート
            var sortedItems = new List<(int Index, JsonNode Item)>();
            for (var i = 0; i < timelineItems.Count; i++)
            {
                var item = timelineItems[i];
                if (item != null)
                {
                    sortedItems.Add((i, item));
                }
            }
            sortedItems.Sort((a, b) =>
            {
                var frameA = a.Item["Frame"]?.GetValue<int>() ?? 0;
                var frameB = b.Item["Frame"]?.GetValue<int>() ?? 0;
                return frameA.CompareTo(frameB);
            });

            // 各無音区間をフレーム単位に変換
            var silentFrameRegions = silentRegions
                .Select(r => (
                    StartFrame: (int)(r.Start.TotalSeconds * fps_val),
                    EndFrame: (int)(r.End.TotalSeconds * fps_val)
                ))
                .OrderBy(r => r.StartFrame)
                .ToList();

            // 各アイテムを処理
            foreach (var (_, item) in sortedItems)
            {
                var type = item["$type"]?.GetValue<string>() ?? "";
                var itemFrame = item["Frame"]?.GetValue<int>() ?? 0;
                var itemLength = item["Length"]?.GetValue<int>() ?? 0;
                var itemEnd = itemFrame + itemLength;

                if (!HasAudioContent(item, type))
                {
                    // 非音声アイテムはフレームオフセットを適用してそのまま追加
                    var cloned = JsonNode.Parse(item.ToJsonString())!;
                    var newFrame = Math.Max(0, itemFrame - frameOffset);
                    cloned["Frame"] = newFrame;
                    newItems.Add(cloned);
                    continue;
                }

                itemsProcessed++;

                // このアイテムに重なる無音区間を見つける
                var overlappingSilences = silentFrameRegions
                    .Where(s => s.StartFrame < itemEnd && s.EndFrame > itemFrame)
                    .ToList();

                if (overlappingSilences.Count == 0)
                {
                    // 無音区間なし => そのままコピー（オフセット適用）
                    var cloned = JsonNode.Parse(item.ToJsonString())!;
                    cloned["Frame"] = Math.Max(0, itemFrame - frameOffset);
                    newItems.Add(cloned);
                    continue;
                }

                // 有音区間を計算（無音区間の補集合）
                var soundRegions = GetSoundRegions(itemFrame, itemEnd, overlappingSilences);

                if (soundRegions.Count == 0)
                {
                    // 全体が無音 => アイテムを削除
                    silentItemsRemoved++;
                    frameOffset += itemLength;
                    totalFramesSaved += itemLength;
                    continue;
                }

                // 有音区間ごとにアイテムを分割
                var isFirst = true;
                foreach (var (regionStart, regionEnd) in soundRegions)
                {
                    var regionLength = regionEnd - regionStart;
                    if (regionLength <= 0) continue;

                    var cloned = JsonNode.Parse(item.ToJsonString())!;

                    // 新しいフレーム位置（オフセット適用）
                    var newFrame = Math.Max(0, regionStart - frameOffset);
                    cloned["Frame"] = newFrame;
                    cloned["Length"] = regionLength;

                    // ContentOffsetを調整（アイテム内の再生開始位置）
                    // ContentOffsetは "HH:MM:SS" or "HH:MM:SS.FFFFFFF" 形式のTimeSpan文字列
                    if (!isFirst || regionStart > itemFrame)
                    {
                        var offsetStr = item["ContentOffset"]?.GetValue<string>() ?? "00:00:00";
                        var existingOffset = TimeSpan.TryParse(offsetStr, out var ts) ? ts : TimeSpan.Zero;
                        var additionalFrames = regionStart - itemFrame;
                        var additionalTime = TimeSpan.FromSeconds(additionalFrames / fps_val);
                        var newOffset = existingOffset + additionalTime;
                        cloned["ContentOffset"] = newOffset.ToString(@"hh\:mm\:ss\.fffffff");
                    }

                    newItems.Add(cloned);

                    if (!isFirst) itemsSplit++;
                    isFirst = false;
                }

                // 無音分のフレームを累計
                var silenceFrames = overlappingSilences
                    .Sum(s =>
                    {
                        var start = Math.Max(s.StartFrame, itemFrame);
                        var end = Math.Min(s.EndFrame, itemEnd);
                        return Math.Max(0, end - start);
                    });
                frameOffset += silenceFrames;
                totalFramesSaved += silenceFrames;
            }

            // アイテムリストを置き換え
            timeline!["Items"] = newItems;
        }

        // 保存
        var options = new JsonSerializerOptions
        {
            WriteIndented = true,
            Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        };
        var outputJson = doc.ToJsonString(options);
        File.WriteAllText(ymmpPath, outputJson);

        var timeSaved = TimeSpan.FromSeconds(totalFramesSaved / fps_val);
        return new CutResult(itemsProcessed, itemsSplit, silentItemsRemoved, timeSaved);
    }

    /// <summary>
    /// 指定範囲内の有音区間（無音区間の補集合）を計算します
    /// </summary>
    private static List<(int Start, int End)> GetSoundRegions(
        int itemStart, int itemEnd,
        List<(int StartFrame, int EndFrame)> silenceRegions)
    {
        var result = new List<(int Start, int End)>();
        var current = itemStart;

        foreach (var (silStart, silEnd) in silenceRegions.OrderBy(s => s.StartFrame))
        {
            var clampedStart = Math.Max(silStart, itemStart);
            var clampedEnd = Math.Min(silEnd, itemEnd);

            if (current < clampedStart)
            {
                result.Add((current, clampedStart));
            }
            current = Math.Max(current, clampedEnd);
        }

        if (current < itemEnd)
        {
            result.Add((current, itemEnd));
        }

        return result;
    }

    /// <summary>
    /// アイテムが音声コンテンツを含むか判定します
    /// </summary>
    private static bool HasAudioContent(JsonNode item, string type)
    {
        // VoiceItem, VideoItem, AudioItem等を判定
        var typeLower = type.ToLowerInvariant();
        if (typeLower.Contains("voiceitem") ||
            typeLower.Contains("audioitem") ||
            typeLower.Contains("videoitem"))
        {
            return true;
        }

        // FilePath が含まれていれば音声/動画ファイルを持つアイテムの可能性
        var filePath = item["FilePath"]?.GetValue<string>();
        if (!string.IsNullOrEmpty(filePath))
        {
            var ext = Path.GetExtension(filePath);
            return MediaExtensions.Supported.Contains(ext);
        }

        return false;
    }

    /// <summary>
    /// プロジェクトのFPSを取得します
    /// </summary>
    private static double? GetProjectFps(JsonNode doc)
    {
        try
        {
            var timelines = doc["Timelines"]?.AsArray();
            if (timelines == null || timelines.Count == 0) return null;

            var videoInfo = timelines[0]?["VideoInfo"];
            if (videoInfo == null) return null;

            var fps = videoInfo["FPS"]?.GetValue<double>();
            return fps;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// ymmpファイルからプロジェクトのFPS値を取得します
    /// </summary>
    public static double GetFps(string ymmpPath)
    {
        var json = File.ReadAllText(ymmpPath);
        var doc = JsonNode.Parse(json);
        if (doc == null) return 30.0;
        return GetProjectFps(doc) ?? 30.0;
    }

    /// <summary>
    /// ymmpファイルに含まれる音声/動画ファイルのパスを取得します
    /// </summary>
    public static List<string> GetMediaFilePaths(string ymmpPath)
    {
        var json = File.ReadAllText(ymmpPath);
        var doc = JsonNode.Parse(json);
        var paths = new List<string>();

        if (doc == null) return paths;

        var timelines = doc["Timelines"]?.AsArray();
        if (timelines == null) return paths;

        foreach (var timeline in timelines)
        {
            var items = timeline?["Items"]?.AsArray();
            if (items == null) continue;

            foreach (var item in items)
            {
                if (item == null) continue;

                var filePath = item["FilePath"]?.GetValue<string>();
                if (!string.IsNullOrEmpty(filePath) && File.Exists(filePath))
                {
                    var ext = Path.GetExtension(filePath);
                    if (MediaExtensions.Supported.Contains(ext))
                    {
                        if (!paths.Contains(filePath))
                            paths.Add(filePath);
                    }
                }
            }
        }

        return paths;
    }
}
