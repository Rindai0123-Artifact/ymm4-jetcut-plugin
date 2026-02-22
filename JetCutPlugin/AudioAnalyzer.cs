using System.IO;
using NAudio.Wave;

namespace JetCutPlugin;

/// <summary>
/// 音声ファイルの音量を解析し、無音区間を検出するクラス
/// </summary>
public class AudioAnalyzer
{
    /// <summary>
    /// 無音区間を表すデータ
    /// </summary>
    public record SilentRegion(TimeSpan Start, TimeSpan End)
    {
        public TimeSpan Duration => End - Start;
    }

    /// <summary>
    /// 音声解析結果
    /// </summary>
    public record AnalysisResult(
        List<SilentRegion> SilentRegions,
        TimeSpan TotalDuration,
        TimeSpan TotalSilenceDuration,
        int SampleRate
    );

    /// <summary>
    /// 音声ファイルを解析し、無音区間を検出します
    /// </summary>
    /// <param name="filePath">音声/動画ファイルのパス</param>
    /// <param name="thresholdDb">無音判定の閾値 (dB)。この値以下を無音とみなす。デフォルト: -40dB</param>
    /// <param name="minSilenceDurationMs">最小無音時間 (ms)。この時間以上続く無音をカット対象にする。デフォルト: 500ms</param>
    /// <param name="marginMs">前後マージン (ms)。カット前後に残す余白。デフォルト: 100ms</param>
    /// <param name="progress">進捗報告用コールバック (0.0～1.0)</param>
    /// <param name="cancellationToken">キャンセルトークン</param>
    /// <returns>解析結果</returns>
    public static async Task<AnalysisResult> AnalyzeAsync(
        string filePath,
        double thresholdDb = -40.0,
        int minSilenceDurationMs = 500,
        int marginMs = 100,
        IProgress<double>? progress = null,
        CancellationToken cancellationToken = default)
    {
        return await Task.Run(() =>
        {
            // MediaFoundationReaderで多くの音声/動画形式に対応
            WaveStream waveStream;
            try
            {
                waveStream = new MediaFoundationReader(filePath);
            }
            catch
            {
                waveStream = new AudioFileReader(filePath);
            }

            using (waveStream)
            {
                var sampleProvider = waveStream.ToSampleProvider();
                var sampleRate = sampleProvider.WaveFormat.SampleRate;
                var channels = sampleProvider.WaveFormat.Channels;
                var totalBytes = waveStream.Length;
                var bytesPerSample = waveStream.WaveFormat.BitsPerSample / 8;
                var totalDuration = waveStream.TotalTime;

                // 解析ウィンドウサイズ: 10ms分のサンプル
                var windowSizeSamples = sampleRate * channels / 100; // 10ms
                var buffer = new float[windowSizeSamples];

                // dBから線形スケールの閾値に変換
                var thresholdLinear = Math.Pow(10.0, thresholdDb / 20.0);

                var silentRegions = new List<SilentRegion>();
                var isSilent = false;
                var silenceStartSample = 0L;
                var currentSample = 0L;
                var totalSamplesEstimate = (double)(totalBytes / bytesPerSample);

                int samplesRead;
                while ((samplesRead = sampleProvider.Read(buffer, 0, buffer.Length)) > 0)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    // RMS (Root Mean Square) 計算
                    var rms = CalculateRms(buffer, samplesRead);

                    if (rms <= thresholdLinear)
                    {
                        // 無音区間に入った
                        if (!isSilent)
                        {
                            silenceStartSample = currentSample;
                            isSilent = true;
                        }
                    }
                    else
                    {
                        // 無音区間が終わった
                        if (isSilent)
                        {
                            AddSilentRegionIfLongEnough(
                                silentRegions, silenceStartSample, currentSample,
                                sampleRate, channels, minSilenceDurationMs, marginMs);
                            isSilent = false;
                        }
                    }

                    currentSample += samplesRead;
                    if (totalSamplesEstimate > 0)
                        progress?.Report(currentSample / totalSamplesEstimate);
                }

                // ファイル末尾が無音の場合
                if (isSilent)
                {
                    var endSample = (long)(totalDuration.TotalSeconds * sampleRate * channels);
                    AddSilentRegionIfLongEnough(
                        silentRegions, silenceStartSample, endSample,
                        sampleRate, channels, minSilenceDurationMs, marginMs);
                }

                var totalSilenceDuration = TimeSpan.FromTicks(
                    silentRegions.Sum(r => r.Duration.Ticks));

                return new AnalysisResult(silentRegions, totalDuration, totalSilenceDuration, sampleRate);
            }
        }, cancellationToken);
    }

    /// <summary>
    /// 無音区間が十分な長さなら追加します
    /// </summary>
    private static void AddSilentRegionIfLongEnough(
        List<SilentRegion> regions, long startSample, long endSample,
        int sampleRate, int channels, int minDurationMs, int marginMs)
    {
        var samplesPerSecond = (double)(sampleRate * channels);
        var startTime = TimeSpan.FromSeconds(startSample / samplesPerSecond);
        var endTime = TimeSpan.FromSeconds(endSample / samplesPerSecond);
        var duration = endTime - startTime;

        if (duration.TotalMilliseconds >= minDurationMs)
        {
            var margin = TimeSpan.FromMilliseconds(marginMs);
            var adjustedStart = startTime + margin;
            var adjustedEnd = endTime - margin;

            if (adjustedEnd > adjustedStart)
            {
                regions.Add(new SilentRegion(adjustedStart, adjustedEnd));
            }
        }
    }

    /// <summary>
    /// サンプルバッファのRMS値を計算します
    /// </summary>
    private static double CalculateRms(float[] buffer, int count)
    {
        if (count == 0) return 0.0;

        var sumOfSquares = 0.0;
        for (var i = 0; i < count; i++)
        {
            sumOfSquares += buffer[i] * (double)buffer[i];
        }
        return Math.Sqrt(sumOfSquares / count);
    }
}
