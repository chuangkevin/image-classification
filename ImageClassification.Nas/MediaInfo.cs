using System.Globalization;
using MetadataExtractor;
using MetadataExtractor.Formats.Exif;
using MetadataExtractor.Formats.QuickTime;

namespace ImageClassification.Nas;

/// <summary>從單一檔案讀到的分類依據。任一欄位可能為 null（讀不到）。</summary>
public sealed record MediaInfo(string? Make, string? Model, DateTime? Taken)
{
    public bool HasCamera => Make is not null || Model is not null;

    private static readonly string[] ExifDateFormats = { "yyyy:MM:dd HH:mm:ss", "yyyy-MM-dd HH:mm:ss", "yyyy:MM:dd HH:mm:ss.fff" };

    public static MediaInfo Read(string path)
    {
        IReadOnlyList<MetadataExtractor.Directory> dirs;
        try
        {
            dirs = ImageMetadataReader.ReadMetadata(path);
        }
        catch
        {
            return new MediaInfo(null, null, null);
        }

        var ifd0 = dirs.OfType<ExifIfd0Directory>().FirstOrDefault();
        var make = Clean(ifd0?.GetDescription(ExifIfd0Directory.TagMake));
        var model = Clean(ifd0?.GetDescription(ExifIfd0Directory.TagModel));

        // 原版 bug：TagDateTime 讀完立刻被 TagDateTimeDigitized 覆蓋。
        // 這裡依可信度依序取：原始拍攝 → 數位化 → IFD0 修改時間。
        DateTime? taken = null;
        foreach (var sub in dirs.OfType<ExifSubIfdDirectory>())
        {
            taken ??= ParseExif(sub.GetDescription(ExifDirectoryBase.TagDateTimeOriginal));
            taken ??= ParseExif(sub.GetDescription(ExifDirectoryBase.TagDateTimeDigitized));
        }
        taken ??= ParseExif(ifd0?.GetDescription(ExifDirectoryBase.TagDateTime));

        // 影片：QuickTime 建立時間（相機多半直接寫當地時間，不做時區換算）。
        if (taken is null)
        {
            var mvhd = dirs.OfType<QuickTimeMovieHeaderDirectory>().FirstOrDefault();
            if (mvhd is not null && mvhd.TryGetDateTime(QuickTimeMovieHeaderDirectory.TagCreated, out var created)
                && created.Year > 1990)
            {
                taken = created;
            }
        }

        if (make is not null && model is not null && model.StartsWith(make, StringComparison.OrdinalIgnoreCase))
        {
            model = Clean(model[make.Length..]);
        }

        return new MediaInfo(make, model, taken);
    }

    private static DateTime? ParseExif(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return null;
        raw = raw.Trim().TrimEnd('\0');
        return DateTime.TryParseExact(raw, ExifDateFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt)
               && dt.Year > 1990
            ? dt
            : null;
    }

    private static string? Clean(string? s)
    {
        if (string.IsNullOrWhiteSpace(s)) return null;
        var chars = s.Trim().TrimEnd('\0').Select(c => c is '/' or '\\' || char.IsControl(c) ? '_' : c).ToArray();
        var v = new string(chars).Trim().Trim('.');
        return v.Length == 0 ? null : v;
    }
}
