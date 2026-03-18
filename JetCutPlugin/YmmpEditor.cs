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
    /// ymmpファイルに対してジェットカットを実行します (v2.0)
    /// </summary>
    /// <param name="ymmpPath">ymmpファイルのパス</param>
    /// <param name="silentRegionsByFile">メディアファイルパス → 無音区間リスト（ファイル時間座標）</param>
    /// <param name="fps">プロジェクトのFPS</param>
    /// <param name="createBackup">バックアップを作成するか</param>
    /// <returns>カット結果</returns>
    public static CutResult ExecuteJetCut(
        string ymmpPath,
        Dictionary<string, List<AudioAnalyzer.SilentRegion>> silentRegionsByFile,
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

        var fps_val = fps > 0 ? fps : (GetProjectFps(doc) ?? 30.0);

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
                if (item != null) sortedItems.Add((i, item));
            }
            sortedItems.Sort((a, b) =>
            {
                var frameA = a.Item["Frame"]?.GetValue<int>() ?? 0;
                var frameB = b.Item["Frame"]?.GetValue<int>() ?? 0;
                return frameA.CompareTo(frameB);
            });

            // 各アイテムを処理
            foreach (var (_, item) in sortedItems)
            {
                var type = item["$type"]?.GetValue<string>() ?? "";
                var itemFrame = item["Frame"]?.GetValue<int>() ?? 0;
                var itemLength = item["Length"]?.GetValue<int>() ?? 0;

                // 非音声アイテム → オフセットのみ適用
                if (!HasAudioContent(item, type))
                {
                    var cloned = JsonNode.Parse(item.ToJsonString())!;
                    cloned["Frame"] = Math.Max(0, itemFrame - frameOffset);
                    newItems.Add(cloned);
                    continue;
                }

                var filePath = item["FilePath"]?.GetValue<string>() ?? "";

                // このファイルにカット対象の無音区間がない → そのまま保持（オフセットのみ適用）
                if (string.IsNullOrEmpty(filePath) ||
                    !silentRegionsByFile.TryGetValue(filePath, out var fileSilences) ||
                    fileSilences.Count == 0)
                {
                    var cloned = JsonNode.Parse(item.ToJsonString())!;
                    cloned["Frame"] = Math.Max(0, itemFrame - frameOffset);
                    newItems.Add(cloned);
                    continue;
                }

                itemsProcessed++;

                // ★ ファイル時間座標でこのアイテムの範囲を算出 ★
                // ContentOffset = このアイテムが元ファイルのどこから再生開始するか
                var offsetStr = item["ContentOffset"]?.GetValue<string>() ?? "00:00:00";
                var contentOffset = TimeSpan.TryParse(offsetStr, out var ts) ? ts : TimeSpan.Zero;
                var fileTimeStart = contentOffset.TotalSeconds;
                var fileTimeEnd = fileTimeStart + itemLength / fps_val;

                // ファイル時間座標で重なる無音区間を検索
                var overlapping = fileSilences
                    .Where(s => s.Start.TotalSeconds < fileTimeEnd && s.End.TotalSeconds > fileTimeStart)
                    .OrderBy(s => s.Start)
                    .ToList();

                if (overlapping.Count == 0)
                {
                    // 重なる無音なし → そのまま保持
                    var cloned = JsonNode.Parse(item.ToJsonString())!;
                    cloned["Frame"] = Math.Max(0, itemFrame - frameOffset);
                    newItems.Add(cloned);
                    continue;
                }

                // ファイル時間座標で有音区間を算出（無音の補集合）
                var soundRegions = new List<(double Start, double End)>();
                var current = fileTimeStart;
                foreach (var sil in overlapping)
                {
                    var silStart = Math.Max(sil.Start.TotalSeconds, fileTimeStart);
                    var silEnd = Math.Min(sil.End.TotalSeconds, fileTimeEnd);
                    if (current < silStart)
                        soundRegions.Add((current, silStart));
                    current = Math.Max(current, silEnd);
                }
                if (current < fileTimeEnd)
                    soundRegions.Add((current, fileTimeEnd));

                if (soundRegions.Count == 0)
                {
                    // 全体が無音 → アイテムを削除
                    silentItemsRemoved++;
                    frameOffset += itemLength;
                    totalFramesSaved += itemLength;
                    continue;
                }

                // 有音区間ごとにアイテムを分割（連続配置）
                var isFirst = true;
                var currentTimelinePos = itemFrame - frameOffset; // このアイテムの開始位置
                foreach (var (sndStart, sndEnd) in soundRegions)
                {
                    var regionLengthFrames = (int)((sndEnd - sndStart) * fps_val);
                    if (regionLengthFrames <= 0) continue;

                    var cloned = JsonNode.Parse(item.ToJsonString())!;

                    // タイムライン上に連続配置（前の区間の直後に置く）
                    cloned["Frame"] = Math.Max(0, currentTimelinePos);
                    cloned["Length"] = regionLengthFrames;

                    // ContentOffset = この分割後アイテムが元ファイルのどこから始まるか
                    var newContentOffset = TimeSpan.FromSeconds(sndStart);
                    cloned["ContentOffset"] = newContentOffset.ToString(@"hh\:mm\:ss\.fffffff");

                    newItems.Add(cloned);
                    currentTimelinePos += regionLengthFrames; // 次の区間は直後に

                    if (!isFirst) itemsSplit++;
                    isFirst = false;
                }

                // 無音分のフレームを累計オフセットに加算
                var silenceFrames = overlapping.Sum(s =>
                {
                    var start = Math.Max(s.Start.TotalSeconds, fileTimeStart);
                    var end = Math.Min(s.End.TotalSeconds, fileTimeEnd);
                    return (int)(Math.Max(0, end - start) * fps_val);
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
        File.WriteAllText(ymmpPath, doc.ToJsonString(options));

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

    /// <summary>
    /// メディアファイルの使用中セグメント情報
    /// </summary>
    public record MediaSegment(TimeSpan Start, TimeSpan End, int TimelineFrame, int TimelineLength);

    /// <summary>
    /// ymmpファイルから各メディアファイルのタイムライン上での使用区間を取得します。
    /// ContentOffset と Length(フレーム) から、元ファイル内のどの区間が使用中かを算出します。
    /// </summary>
    public static Dictionary<string, List<MediaSegment>> GetMediaSegments(string ymmpPath)
    {
        var json = File.ReadAllText(ymmpPath);
        var doc = JsonNode.Parse(json);
        var result = new Dictionary<string, List<MediaSegment>>(StringComparer.OrdinalIgnoreCase);

        if (doc == null) return result;

        var fps = GetProjectFps(doc) ?? 30.0;
        var timelines = doc["Timelines"]?.AsArray();
        if (timelines == null) return result;

        foreach (var timeline in timelines)
        {
            var items = timeline?["Items"]?.AsArray();
            if (items == null) continue;

            foreach (var item in items)
            {
                if (item == null) continue;

                var filePath = item["FilePath"]?.GetValue<string>();
                if (string.IsNullOrEmpty(filePath)) continue;

                var ext = Path.GetExtension(filePath);
                if (!MediaExtensions.Supported.Contains(ext)) continue;

                var type = item["$type"]?.GetValue<string>() ?? "";
                if (!HasAudioContent(item, type)) continue;

                // ContentOffset: "HH:MM:SS" or "HH:MM:SS.FFFFFFF" 形式
                var offsetStr = item["ContentOffset"]?.GetValue<string>() ?? "00:00:00";
                var contentOffset = TimeSpan.TryParse(offsetStr, out var ts) ? ts : TimeSpan.Zero;
                var lengthFrames = item["Length"]?.GetValue<int>() ?? 0;
                var frame = item["Frame"]?.GetValue<int>() ?? 0;

                var segStart = contentOffset;
                var segEnd = contentOffset + TimeSpan.FromSeconds(lengthFrames / fps);

                if (!result.ContainsKey(filePath))
                    result[filePath] = new List<MediaSegment>();

                result[filePath].Add(new MediaSegment(segStart, segEnd, frame, lengthFrames));
            }
        }

        return result;
    }
}
