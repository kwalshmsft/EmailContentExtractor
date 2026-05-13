using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using HtmlAgilityPack;

namespace EmailContentExtractor.Services;

/// <summary>
/// Processes Iris Content XML export files, adding localized Variant elements
/// for each translated email version.
/// </summary>
public class IrisContentImportService
{
    private readonly FileStorageService _fileStorage;

    private static readonly HashSet<string> SkipElements = new(StringComparer.OrdinalIgnoreCase)
    {
        "script", "style", "noscript", "head", "meta", "link"
    };

    private static readonly HashSet<string> BlockElements = new(StringComparer.OrdinalIgnoreCase)
    {
        "p", "div", "tr", "h1", "h2", "h3", "h4", "h5", "h6",
        "blockquote", "section", "article", "header", "footer", "nav",
        "table", "thead", "tbody", "tfoot", "ul", "ol", "dd", "dt"
    };

    public IrisContentImportService(FileStorageService fileStorage)
    {
        _fileStorage = fileStorage;
    }

    /// <summary>
    /// Parses an Iris Content XML export file and returns metadata about its content items.
    /// </summary>
    public IrisContentInfo ParseExport(string xmlContent)
    {
        var doc = XDocument.Parse(xmlContent);
        var root = doc.Root ?? throw new InvalidOperationException("Invalid XML: no root element.");

        var items = root.Elements("ExportedContentItem").Select(item =>
        {
            var existingVariants = item.Elements("Variant")
                .Select(v => v.Attribute("variantCulture")?.Value ?? "unknown")
                .ToList();

            return new IrisContentItemInfo
            {
                Name = item.Attribute("name")?.Value ?? "unknown",
                Id = item.Attribute("id")?.Value ?? "",
                ExistingVariants = existingVariants
            };
        }).ToList();

        return new IrisContentInfo { Items = items };
    }

    /// <summary>
    /// Generates a ZIP archive with locale-named subfolders, each containing the XML file
    /// with that locale's variant content. This matches the Iris Content import folder structure:
    ///   root/InvariantCulture/file.xml
    ///   root/fr-FR/file.xml
    ///   root/de-DE/file.xml
    /// </summary>
    public IrisImportResult GenerateImportZip(string xmlContent, string xmlFileName, string sourceFileName)
    {
        var doc = XDocument.Parse(xmlContent);
        var root = doc.Root ?? throw new InvalidOperationException("Invalid XML: no root element.");

        var baseName = FileStorageService.ParseBaseName(sourceFileName);
        if (baseName == null)
            throw new InvalidOperationException($"Could not parse base name from \"{sourceFileName}\".");

        var sourceHtml = _fileStorage.GetSourceContent(sourceFileName);
        if (sourceHtml == null)
            throw new InvalidOperationException($"Source file \"{sourceFileName}\" not found.");

        // Collect generated localized files
        var localizedFiles = new List<(string locale, string html)>();

        var generatedFiles = _fileStorage.GetAllGenerated()
            .Where(g => g.BaseName != null &&
                        g.BaseName.Equals(baseName, StringComparison.OrdinalIgnoreCase))
            .ToList();

        foreach (var gen in generatedFiles)
        {
            var genLocale = gen.Locale;
            if (genLocale == null) continue;
            var genHtml = _fileStorage.GetGeneratedContent(gen.FileName);
            if (genHtml == null) continue;
            localizedFiles.Add((genLocale, genHtml));
        }

        var variants = new List<string>();

        // Build per-locale XML files (skip InvariantCulture — import ignores it)
        var localeXmlFiles = new List<(string folder, string xml)>();

        // Each generated locale gets its own subfolder
        foreach (var (locale, html) in localizedFiles)
        {
            localeXmlFiles.Add((locale, BuildVariantXml(doc, html)));
            variants.Add(locale);
        }

        // Create ZIP with folder structure
        using var ms = new MemoryStream();
        using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var (folder, xml) in localeXmlFiles)
            {
                var entryPath = $"{folder}/{xmlFileName}";
                var entry = archive.CreateEntry(entryPath, CompressionLevel.Optimal);
                using var entryStream = entry.Open();
                using var writer = new StreamWriter(entryStream, new UTF8Encoding(false));
                writer.Write(xml);
            }
        }

        return new IrisImportResult
        {
            ZipBytes = ms.ToArray(),
            Variants = variants
        };
    }

    /// <summary>
    /// Creates a copy of the XML with the Variant content replaced.
    /// The variantCulture stays as InvariantCulture — the folder name determines the target locale.
    /// </summary>
    private string BuildVariantXml(XDocument originalDoc, string html)
    {
        var doc = new XDocument(originalDoc);
        var root = doc.Root!;

        var subject = HtmlHelper.ExtractTitleText(html);
        var plainText = ConvertHtmlToPlainText(html);

        foreach (var item in root.Elements("ExportedContentItem"))
        {
            // Replace content in the existing Variant (keep original variantCulture)
            var variant = item.Elements("Variant").FirstOrDefault();
            if (variant != null)
            {
                variant.RemoveNodes();
                variant.Add(
                    new XElement("Field",
                        new XAttribute("name", "Subject/Content"),
                        new XAttribute("type", "String"),
                        subject),
                    new XElement("Field",
                        new XAttribute("name", "PlainTextBody/Content"),
                        new XAttribute("type", "String"),
                        plainText),
                    new XElement("Field",
                        new XAttribute("name", "HtmlBodyFormat/Content"),
                        new XAttribute("type", "RichContent"),
                        new XCData(html))
                );
            }
        }

        return doc.Declaration != null
            ? doc.Declaration.ToString() + "\n" + doc.Root!.ToString()
            : doc.Root!.ToString();
    }

    /// <summary>
    /// Converts HTML to plain text suitable for email PlainTextBody.
    /// Preserves link URLs, converts block elements to newlines, strips all tags.
    /// Word-wraps lines near 72 characters without exceeding 75.
    /// </summary>
    public static string ConvertHtmlToPlainText(string html)
    {
        var doc = new HtmlDocument();
        doc.LoadHtml(html);

        var sb = new StringBuilder();
        WalkNode(doc.DocumentNode, sb);

        // Clean up whitespace
        var text = sb.ToString();
        // Remove bullet characters entirely
        text = text.Replace("•", "").Replace("·", "");
        text = text.Replace("\u00A0", " ");   // non-breaking space
        text = text.Replace("\u200B", "");     // zero-width space
        text = text.Replace("\u200C", "");     // zero-width non-joiner
        text = text.Replace("\u200D", "");     // zero-width joiner
        text = text.Replace("\u2018", "'").Replace("\u2019", "'");   // smart quotes
        text = text.Replace("\u201C", "\"").Replace("\u201D", "\""); // smart double quotes
        text = text.Replace("\u2013", "-").Replace("\u2014", "-");   // en/em dash
        text = text.Replace("\u2026", "...");  // ellipsis
        // Left-justify all lines (no leading whitespace)
        text = Regex.Replace(text, @"^[ \t]+", "", RegexOptions.Multiline);
        // Collapse 3+ newlines to 2 (one blank line between paragraphs max)
        text = Regex.Replace(text, @"(\r?\n){3,}", "\n\n");
        // Trim trailing whitespace on each line
        text = Regex.Replace(text, @"[ \t]+$", "", RegexOptions.Multiline);
        text = text.Trim();

        // Word-wrap lines near 72 characters
        return WordWrap(text, 72, 75);
    }

    /// <summary>
    /// Wraps text at word boundaries, targeting targetWidth but never exceeding maxWidth.
    /// </summary>
    private static string WordWrap(string text, int targetWidth, int maxWidth)
    {
        var result = new StringBuilder();
        foreach (var line in text.Split('\n'))
        {
            if (line.Length <= maxWidth)
            {
                result.AppendLine(line);
                continue;
            }

            var remaining = line;
            while (remaining.Length > maxWidth)
            {
                // Find the last space at or before targetWidth
                var breakAt = remaining.LastIndexOf(' ', Math.Min(targetWidth, remaining.Length - 1));
                if (breakAt <= 0)
                {
                    // No space found before target; find the first space before maxWidth
                    breakAt = remaining.LastIndexOf(' ', Math.Min(maxWidth, remaining.Length - 1));
                }
                if (breakAt <= 0)
                {
                    // No space at all before maxWidth; hard break
                    breakAt = maxWidth;
                }

                result.AppendLine(remaining[..breakAt].TrimEnd());
                remaining = remaining[breakAt..].TrimStart();
            }
            result.AppendLine(remaining);
        }

        // Remove trailing newline added by last AppendLine
        if (result.Length >= Environment.NewLine.Length)
            result.Length -= Environment.NewLine.Length;

        return result.ToString();
    }

    private static void WalkNode(HtmlNode node, StringBuilder sb)
    {
        if (node.NodeType == HtmlNodeType.Comment)
            return;

        if (node.NodeType == HtmlNodeType.Text)
        {
            var text = HtmlEntity.DeEntitize(node.InnerText);
            // Collapse internal whitespace but preserve single spaces
            text = Regex.Replace(text, @"[ \t]+", " ");
            // Don't add pure-whitespace text that's just newlines
            if (!string.IsNullOrWhiteSpace(text))
                sb.Append(text);
            return;
        }

        if (node.NodeType != HtmlNodeType.Element)
        {
            foreach (var child in node.ChildNodes)
                WalkNode(child, sb);
            return;
        }

        var tag = node.Name.ToLowerInvariant();

        // Skip invisible elements
        if (SkipElements.Contains(tag))
            return;

        // Skip hidden elements
        var style = node.GetAttributeValue("style", "");
        if (style.Contains("display:none", StringComparison.OrdinalIgnoreCase) ||
            style.Contains("display: none", StringComparison.OrdinalIgnoreCase) ||
            style.Contains("visibility:hidden", StringComparison.OrdinalIgnoreCase) ||
            style.Contains("visibility: hidden", StringComparison.OrdinalIgnoreCase) ||
            style.Contains("mso-hide:all", StringComparison.OrdinalIgnoreCase))
            return;

        if (node.GetAttributeValue("aria-hidden", "") == "true")
            return;

        // Handle specific tags
        switch (tag)
        {
            case "br":
                sb.AppendLine();
                return;

            case "hr":
                sb.AppendLine();
                return;

            case "li":
                // Just output the text content, no bullet prefix
                foreach (var child in node.ChildNodes)
                    WalkNode(child, sb);
                sb.AppendLine();
                return;

            case "a":
                var href = node.GetAttributeValue("href", "");
                // Check if link wraps only an image
                var hasOnlyImg = node.ChildNodes.All(c =>
                    c.NodeType == HtmlNodeType.Element && c.Name.Equals("img", StringComparison.OrdinalIgnoreCase)
                    || c.NodeType == HtmlNodeType.Text && string.IsNullOrWhiteSpace(c.InnerText));
                if (hasOnlyImg)
                {
                    // Image link: just output the URL on its own line
                    if (!string.IsNullOrEmpty(href))
                    {
                        sb.AppendLine();
                        sb.Append(href);
                    }
                    return;
                }
                // Text link: inline with URL in parentheses
                var linkText = new StringBuilder();
                foreach (var child in node.ChildNodes)
                {
                    if (child.NodeType == HtmlNodeType.Element && child.Name.Equals("img", StringComparison.OrdinalIgnoreCase))
                        continue;
                    WalkNode(child, linkText);
                }
                var lt = linkText.ToString().Trim();
                if (!string.IsNullOrEmpty(lt) && !string.IsNullOrEmpty(href) && lt != href)
                    sb.Append($"{lt} ({href})");
                else if (!string.IsNullOrEmpty(lt))
                    sb.Append(lt);
                else if (!string.IsNullOrEmpty(href))
                    sb.Append(href);
                return;

            case "img":
                // Skip images — no alt tags in plain text
                return;
        }

        // Block-level elements get newlines
        bool isBlock = BlockElements.Contains(tag);
        if (isBlock)
            sb.AppendLine();

        foreach (var child in node.ChildNodes)
            WalkNode(child, sb);

        if (isBlock)
            sb.AppendLine();
    }
}

/// <summary>
/// Shared HTML helper methods.
/// </summary>
public static class HtmlHelper
{
    /// <summary>
    /// Extracts the text content of the title tag from HTML.
    /// </summary>
    public static string ExtractTitleText(string html)
    {
        var doc = new HtmlDocument();
        doc.LoadHtml(html);
        var titleNode = doc.DocumentNode.SelectSingleNode("//title");
        return titleNode != null ? HtmlEntity.DeEntitize(titleNode.InnerText).Trim() : "";
    }
}

public class IrisContentInfo
{
    public List<IrisContentItemInfo> Items { get; set; } = new();
}

public class IrisContentItemInfo
{
    public string Name { get; set; } = string.Empty;
    public string Id { get; set; } = string.Empty;
    public List<string> ExistingVariants { get; set; } = new();
}

public class IrisImportResult
{
    public byte[] ZipBytes { get; set; } = Array.Empty<byte>();
    public List<string> Variants { get; set; } = new();
}
