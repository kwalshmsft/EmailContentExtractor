using System.Text.RegularExpressions;

namespace EmailContentExtractor.Services;

public class FileStorageService
{
    private readonly Dictionary<string, StoredFile> _sources = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, StoredFile> _generated = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, byte[]> _csv = new(StringComparer.OrdinalIgnoreCase);

    private static readonly Regex LocalePattern = new(
        @"_([a-zA-Z]{2,3}(?:-[a-zA-Z]{2,4})?)(_strings)?(\s*\(\d+\))?\.[^.]+$",
        RegexOptions.Compiled);

    /// <summary>
    /// Parses locale from a filename like "email_en-au.html" → "en-AU"
    /// or "email_en.html" → "en".
    /// Returns null if no locale pattern is found.
    /// </summary>
    public static string? ParseLocale(string fileName)
    {
        var match = LocalePattern.Match(fileName);
        if (!match.Success) return null;

        var raw = match.Groups[1].Value;
        var parts = raw.Split('-');
        if (parts.Length == 2)
            return $"{parts[0].ToLowerInvariant()}-{parts[1].ToUpperInvariant()}";
        return parts[0].ToLowerInvariant();
    }

    /// <summary>
    /// Extracts the base name from a filename (everything before the locale).
    /// e.g. "email_en-au.html" → "email"
    /// </summary>
    public static string? ParseBaseName(string fileName)
    {
        var match = LocalePattern.Match(fileName);
        if (!match.Success) return null;

        var nameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
        // Strip trailing browser duplicate suffix like " (2)"
        nameWithoutExt = Regex.Replace(nameWithoutExt, @"\s*\(\d+\)$", "");
        // Strip trailing "_strings" if present (from downloaded CSV filenames)
        if (nameWithoutExt.EndsWith("_strings", StringComparison.OrdinalIgnoreCase))
            nameWithoutExt = nameWithoutExt[..^"_strings".Length];

        var localeWithUnderscore = "_" + match.Groups[1].Value;
        var idx = nameWithoutExt.LastIndexOf(localeWithUnderscore, StringComparison.OrdinalIgnoreCase);
        return idx >= 0 ? nameWithoutExt[..idx] : null;
    }

    public Task<string> StoreSourceAsync(string fileName, string content)
    {
        _sources[fileName] = new StoredFile(fileName, content, DateTime.UtcNow);
        return Task.FromResult(fileName);
    }

    public string? GetSourceContent(string fileName)
    {
        return _sources.TryGetValue(fileName, out var f) ? f.Content : null;
    }

    public Task StoreCsvAsync(string fileName, byte[] csvBytes)
    {
        _csv[fileName] = csvBytes;
        return Task.CompletedTask;
    }

    public byte[]? GetCsvBytes(string fileName)
    {
        return _csv.TryGetValue(fileName, out var bytes) ? bytes : null;
    }

    public Task StoreGeneratedAsync(string fileName, string content)
    {
        _generated[fileName] = new StoredFile(fileName, content, DateTime.UtcNow);
        return Task.CompletedTask;
    }

    public string? GetGeneratedContent(string fileName)
    {
        return _generated.TryGetValue(fileName, out var f) ? f.Content : null;
    }

    public byte[]? GetGeneratedBytes(string fileName)
    {
        return _generated.TryGetValue(fileName, out var f)
            ? System.Text.Encoding.UTF8.GetBytes(f.Content)
            : null;
    }

    public List<StoredFileInfo> GetAllSources()
    {
        return _sources.Values
            .Select(f => ToFileInfo(f))
            .OrderByDescending(f => f.CreatedUtc)
            .ToList();
    }

    public List<StoredFileInfo> GetAllGenerated()
    {
        return _generated.Values
            .Select(f => ToFileInfo(f))
            .OrderByDescending(f => f.CreatedUtc)
            .ToList();
    }

    public StoredFileInfo? FindSourceByBaseName(string baseName)
    {
        return GetAllSources()
            .FirstOrDefault(s => s.BaseName != null &&
                s.BaseName.Equals(baseName, StringComparison.OrdinalIgnoreCase));
    }

    public void DeleteSource(string fileName)
    {
        _sources.Remove(fileName);
        var csvName = Path.GetFileNameWithoutExtension(fileName) + "_strings.csv";
        _csv.Remove(csvName);
    }

    public void DeleteGenerated(string fileName)
    {
        _generated.Remove(fileName);
    }

    private static StoredFileInfo ToFileInfo(StoredFile f)
    {
        return new StoredFileInfo
        {
            FileName = f.FileName,
            Locale = ParseLocale(f.FileName),
            BaseName = ParseBaseName(f.FileName),
            CreatedUtc = f.CreatedUtc,
            SizeBytes = System.Text.Encoding.UTF8.GetByteCount(f.Content)
        };
    }

    private record StoredFile(string FileName, string Content, DateTime CreatedUtc);
}

public class StoredFileInfo
{
    public string FileName { get; set; } = string.Empty;
    public string? Locale { get; set; }
    public string? BaseName { get; set; }
    public DateTime CreatedUtc { get; set; }
    public long SizeBytes { get; set; }
}
