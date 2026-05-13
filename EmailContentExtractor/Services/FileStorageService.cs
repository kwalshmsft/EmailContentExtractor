using System.Text.RegularExpressions;

namespace EmailContentExtractor.Services;

public class FileStorageService
{
    private readonly string _sourcesDir;
    private readonly string _generatedDir;
    private readonly string _csvDir;
    private static readonly Regex LocalePattern = new(
        @"_([a-zA-Z]{2}-[a-zA-Z]{2})(_strings)?(\s*\(\d+\))?\.[^.]+$",
        RegexOptions.Compiled);

    public FileStorageService(IWebHostEnvironment env)
    {
        var dataDir = Path.Combine(env.ContentRootPath, "Data");
        _sourcesDir = Path.Combine(dataDir, "sources");
        _generatedDir = Path.Combine(dataDir, "generated");
        _csvDir = Path.Combine(dataDir, "csv");

        Directory.CreateDirectory(_sourcesDir);
        Directory.CreateDirectory(_generatedDir);
        Directory.CreateDirectory(_csvDir);
    }

    /// <summary>
    /// Parses locale from a filename like "email_K_en-au.html" → "en-AU".
    /// Returns null if no locale pattern is found.
    /// </summary>
    public static string? ParseLocale(string fileName)
    {
        var match = LocalePattern.Match(fileName);
        if (!match.Success) return null;

        var raw = match.Groups[1].Value; // e.g. "en-au"
        var parts = raw.Split('-');
        return $"{parts[0].ToLowerInvariant()}-{parts[1].ToUpperInvariant()}";
    }

    /// <summary>
    /// Extracts the base name from a filename (everything before the locale).
    /// e.g. "email_K_en-au.html" → "email_K"
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

    public async Task<string> StoreSourceAsync(string fileName, string content)
    {
        var path = Path.Combine(_sourcesDir, fileName);
        await File.WriteAllTextAsync(path, content);
        return fileName;
    }

    public string? GetSourceContent(string fileName)
    {
        var path = Path.Combine(_sourcesDir, fileName);
        return File.Exists(path) ? File.ReadAllText(path) : null;
    }

    public async Task StoreCsvAsync(string fileName, byte[] csvBytes)
    {
        var path = Path.Combine(_csvDir, fileName);
        await File.WriteAllBytesAsync(path, csvBytes);
    }

    public byte[]? GetCsvBytes(string fileName)
    {
        var path = Path.Combine(_csvDir, fileName);
        return File.Exists(path) ? File.ReadAllBytes(path) : null;
    }

    public async Task StoreGeneratedAsync(string fileName, string content)
    {
        var path = Path.Combine(_generatedDir, fileName);
        await File.WriteAllTextAsync(path, content);
    }

    public string? GetGeneratedContent(string fileName)
    {
        var path = Path.Combine(_generatedDir, fileName);
        return File.Exists(path) ? File.ReadAllText(path) : null;
    }

    public byte[]? GetGeneratedBytes(string fileName)
    {
        var path = Path.Combine(_generatedDir, fileName);
        return File.Exists(path) ? File.ReadAllBytes(path) : null;
    }

    /// <summary>
    /// Returns all stored source files with their parsed info.
    /// </summary>
    public List<StoredFileInfo> GetAllSources()
    {
        return GetFileInfos(_sourcesDir);
    }

    /// <summary>
    /// Returns all generated files with their parsed info.
    /// </summary>
    public List<StoredFileInfo> GetAllGenerated()
    {
        return GetFileInfos(_generatedDir);
    }

    /// <summary>
    /// Finds a source file matching the given base name.
    /// </summary>
    public StoredFileInfo? FindSourceByBaseName(string baseName)
    {
        return GetAllSources()
            .FirstOrDefault(s => s.BaseName != null &&
                s.BaseName.Equals(baseName, StringComparison.OrdinalIgnoreCase));
    }

    public void DeleteSource(string fileName)
    {
        var path = Path.Combine(_sourcesDir, fileName);
        if (File.Exists(path)) File.Delete(path);

        // Also delete associated CSV
        var csvName = Path.GetFileNameWithoutExtension(fileName) + "_strings.csv";
        var csvPath = Path.Combine(_csvDir, csvName);
        if (File.Exists(csvPath)) File.Delete(csvPath);
    }

    public void DeleteGenerated(string fileName)
    {
        var path = Path.Combine(_generatedDir, fileName);
        if (File.Exists(path)) File.Delete(path);
    }

    private static List<StoredFileInfo> GetFileInfos(string directory)
    {
        if (!Directory.Exists(directory))
            return new List<StoredFileInfo>();

        return Directory.GetFiles(directory)
            .Select(path =>
            {
                var name = Path.GetFileName(path);
                return new StoredFileInfo
                {
                    FileName = name,
                    Locale = ParseLocale(name),
                    BaseName = ParseBaseName(name),
                    CreatedUtc = File.GetCreationTimeUtc(path),
                    SizeBytes = new FileInfo(path).Length
                };
            })
            .OrderByDescending(f => f.CreatedUtc)
            .ToList();
    }
}

public class StoredFileInfo
{
    public string FileName { get; set; } = string.Empty;
    public string? Locale { get; set; }
    public string? BaseName { get; set; }
    public DateTime CreatedUtc { get; set; }
    public long SizeBytes { get; set; }
}
