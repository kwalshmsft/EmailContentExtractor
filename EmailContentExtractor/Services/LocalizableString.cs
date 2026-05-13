namespace EmailContentExtractor.Services;

public class LocalizableString
{
    public string Key { get; set; } = string.Empty;
    public string Type { get; set; } = string.Empty; // "text" or "attribute"
    public string Original { get; set; } = string.Empty;
    public string Path { get; set; } = string.Empty;
    public string Attribute { get; set; } = string.Empty; // e.g. "alt", "title"
}
