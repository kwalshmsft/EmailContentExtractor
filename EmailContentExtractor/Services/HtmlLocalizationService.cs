using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using MiniExcelLibs;
using CsvHelper;
using CsvHelper.Configuration;
using HtmlAgilityPack;

namespace EmailContentExtractor.Services;

public class HtmlLocalizationService
{
    private static readonly HashSet<string> SkipElements = new(StringComparer.OrdinalIgnoreCase)
    {
        "script", "style", "noscript"
    };

    private static readonly HashSet<string> LocalizableAttributes = new(StringComparer.OrdinalIgnoreCase)
    {
        "alt", "title"
    };

    private static readonly HashSet<string> InlineElements = new(StringComparer.OrdinalIgnoreCase)
    {
        "a", "abbr", "b", "bdi", "bdo", "br", "cite", "code", "data", "dfn",
        "em", "i", "kbd", "mark", "q", "rp", "rt", "ruby", "s", "samp",
        "small", "span", "strong", "sub", "sup", "time", "u", "var", "wbr",
        "font", "img"
    };

    /// <summary>
    /// Extracts all localizable strings from an HTML document.
    /// </summary>
    public List<LocalizableString> ExtractStrings(string html)
    {
        var doc = new HtmlDocument();
        doc.LoadHtml(html);

        var results = new List<(LocalizableString info, HtmlNode node)>();
        int counter = 0;
        int subjectCounter = 0;

        TraverseNodes(doc.DocumentNode, results, ref counter, ref subjectCounter);

        return results.Select(r => r.info).ToList();
    }

    /// <summary>
    /// Shared traversal used by both extraction and localized HTML generation.
    /// Coalesces inline content (links, bold, etc.) into single strings and
    /// filters out text that contains no letter characters.
    /// </summary>
    private void TraverseNodes(HtmlNode node, List<(LocalizableString info, HtmlNode node)> results, ref int counter, ref int subjectCounter)
    {
        if (node.NodeType == HtmlNodeType.Element && SkipElements.Contains(node.Name))
            return;

        if (IsHiddenElement(node))
            return;

        // Extract localizable attributes (alt, title)
        if (node.NodeType == HtmlNodeType.Element && node.HasAttributes)
        {
            foreach (var attrName in LocalizableAttributes)
            {
                var attr = node.Attributes[attrName];
                if (attr != null && !string.IsNullOrWhiteSpace(attr.Value))
                {
                    counter++;
                    results.Add((new LocalizableString
                    {
                        Key = $"attr_{counter}_{attrName}",
                        Type = "attribute",
                        Original = HtmlEntity.DeEntitize(attr.Value),
                        Path = node.XPath,
                        Attribute = attrName
                    }, node));
                }
            }
        }

        // If this element's children are all text/inline elements, extract as one unit
        if (HasInlineContent(node))
        {
            var innerHtml = node.InnerHtml.Trim();
            var innerText = HtmlEntity.DeEntitize(node.InnerText);
            if (ContainsLetters(innerText))
            {
                counter++;
                var key = IsInsideTitleElement(node)
                    ? NextSubjectKey(ref subjectCounter)
                    : $"html_{counter}";
                results.Add((new LocalizableString
                {
                    Key = key,
                    Type = "html",
                    Original = innerHtml,
                    Path = node.XPath,
                    Attribute = string.Empty
                }, node));
            }
            // Don't recurse — inline children are part of the coalesced string
            return;
        }

        // Extract text nodes that contain at least one letter
        if (node.NodeType == HtmlNodeType.Text)
        {
            var text = HtmlEntity.DeEntitize(node.InnerText);
            if (!string.IsNullOrWhiteSpace(text) && ContainsLetters(text.Trim()))
            {
                counter++;
                var key = IsInsideTitleElement(node.ParentNode)
                    ? NextSubjectKey(ref subjectCounter)
                    : $"text_{counter}";
                results.Add((new LocalizableString
                {
                    Key = key,
                    Type = "text",
                    Original = text.Trim(),
                    Path = node.XPath,
                    Attribute = string.Empty
                }, node));
            }
        }

        // Recurse into children
        foreach (var child in node.ChildNodes)
        {
            TraverseNodes(child, results, ref counter, ref subjectCounter);
        }
    }

    private static bool IsInsideTitleElement(HtmlNode? node)
    {
        return node is { NodeType: HtmlNodeType.Element } &&
               node.Name.Equals("title", StringComparison.OrdinalIgnoreCase);
    }

    private static string NextSubjectKey(ref int subjectCounter)
    {
        subjectCounter++;
        return subjectCounter == 1 ? "subject" : $"subject_{subjectCounter}";
    }

    private static bool ContainsLetters(string text)
    {
        foreach (var c in text)
        {
            if (char.IsLetter(c)) return true;
        }
        return false;
    }

    /// <summary>
    /// Returns true if a node's children are all text nodes or inline elements
    /// (with at least one inline element). Comments block coalescing to
    /// protect email conditional comments.
    /// </summary>
    private static bool HasInlineContent(HtmlNode node)
    {
        if (node.NodeType != HtmlNodeType.Element || !node.HasChildNodes)
            return false;

        bool hasInlineElement = false;

        foreach (var child in node.ChildNodes)
        {
            switch (child.NodeType)
            {
                case HtmlNodeType.Text:
                    continue;
                case HtmlNodeType.Comment:
                    return false;
                case HtmlNodeType.Element when InlineElements.Contains(child.Name):
                    hasInlineElement = true;
                    continue;
                default:
                    return false;
            }
        }

        return hasInlineElement;
    }

    private static bool IsHiddenElement(HtmlNode node)
    {
        if (node.NodeType != HtmlNodeType.Element)
            return false;

        var style = node.GetAttributeValue("style", "");
        if (style.Contains("display:none", StringComparison.OrdinalIgnoreCase) ||
            style.Contains("display: none", StringComparison.OrdinalIgnoreCase) ||
            style.Contains("visibility:hidden", StringComparison.OrdinalIgnoreCase) ||
            style.Contains("visibility: hidden", StringComparison.OrdinalIgnoreCase))
            return false; // Don't skip — email preheaders often use this but may still need translation

        if (node.GetAttributeValue("aria-hidden", "") == "true")
            return false; // Same reasoning for email

        return false;
    }

    /// <summary>
    /// Exports extracted strings to CSV with UTF-8 BOM for Excel compatibility.
    /// Includes an empty "Translation" column for the user to fill in.
    /// </summary>
    public byte[] ExportToCsv(List<LocalizableString> strings)
    {
        using var ms = new MemoryStream();
        // Write UTF-8 BOM
        ms.Write(Encoding.UTF8.GetPreamble());

        using var writer = new StreamWriter(ms, new UTF8Encoding(false), leaveOpen: true);
        using var csv = new CsvWriter(writer, new CsvConfiguration(CultureInfo.InvariantCulture)
        {
            ShouldQuote = _ => true // Quote all fields for safety
        });

        // Write header with Translation column
        csv.WriteField("Key");
        csv.WriteField("Type");
        csv.WriteField("Original");
        csv.WriteField("Translation");
        csv.WriteField("Path");
        csv.WriteField("Attribute");
        csv.NextRecord();

        foreach (var s in strings)
        {
            csv.WriteField(s.Key);
            csv.WriteField(s.Type);
            csv.WriteField(s.Original);
            csv.WriteField(""); // Translation — user fills this in
            csv.WriteField(s.Path);
            csv.WriteField(s.Attribute);
            csv.NextRecord();
        }

        writer.Flush();
        return ms.ToArray();
    }

    /// <summary>
    /// Parses a translated CSV where the user has filled in the "Translation" column.
    /// Returns the original LocalizableStrings and a key→translation dictionary.
    /// </summary>
    public (List<LocalizableString> originals, Dictionary<string, string> translations)
        ParseSimpleTranslatedCsv(byte[] csvBytes)
    {
        using var ms = new MemoryStream(csvBytes);
        using var reader = new StreamReader(ms, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
        using var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture)
        {
            HeaderValidated = null,
            MissingFieldFound = null
        });

        csv.Read();
        csv.ReadHeader();

        var originals = new List<LocalizableString>();
        var translations = new Dictionary<string, string>();

        while (csv.Read())
        {
            var key = csv.GetField("Key") ?? "";
            var original = new LocalizableString
            {
                Key = key,
                Type = csv.GetField("Type") ?? "",
                Original = csv.GetField("Original") ?? "",
                Path = csv.GetField("Path") ?? "",
                Attribute = csv.GetField("Attribute") ?? ""
            };
            originals.Add(original);

            var translation = csv.GetField("Translation");
            if (!string.IsNullOrWhiteSpace(translation))
            {
                translations[key] = translation;
            }
        }

        return (originals, translations);
    }

    /// <summary>
    /// Reads a translated CSV and returns a dictionary of language -> list of (key, translation).
    /// Also returns the list of LocalizableString from the CSV for validation.
    /// </summary>
    public (List<string> languages, Dictionary<string, Dictionary<string, string>> translations, List<LocalizableString> originals)
        ParseTranslatedCsv(byte[] csvBytes)
    {
        using var ms = new MemoryStream(csvBytes);
        using var reader = new StreamReader(ms, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
        using var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture)
        {
            HeaderValidated = null,
            MissingFieldFound = null
        });

        csv.Read();
        csv.ReadHeader();

        var headerRecord = csv.HeaderRecord!;
        var knownHeaders = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Key", "Type", "Original", "Path", "Attribute"
        };

        var languages = headerRecord
            .Where(h => !knownHeaders.Contains(h))
            .ToList();

        var translations = new Dictionary<string, Dictionary<string, string>>();
        foreach (var lang in languages)
        {
            translations[lang] = new Dictionary<string, string>();
        }

        var originals = new List<LocalizableString>();

        while (csv.Read())
        {
            var key = csv.GetField("Key") ?? "";
            var original = new LocalizableString
            {
                Key = key,
                Type = csv.GetField("Type") ?? "",
                Original = csv.GetField("Original") ?? "",
                Path = csv.GetField("Path") ?? "",
                Attribute = csv.GetField("Attribute") ?? ""
            };
            originals.Add(original);

            foreach (var lang in languages)
            {
                var value = csv.GetField(lang);
                if (!string.IsNullOrWhiteSpace(value))
                {
                    translations[lang][key] = value;
                }
            }
        }

        return (languages, translations, originals);
    }

    /// <summary>
    /// Generates localized HTML for a single language.
    /// Uses the same traversal order as extraction to match keys to nodes.
    /// Returns the localized HTML and a list of warnings.
    /// </summary>
    public (string html, List<string> warnings) GenerateLocalizedHtml(
        string originalHtml,
        List<LocalizableString> originals,
        Dictionary<string, string> translations)
    {
        var doc = new HtmlDocument();
        doc.OptionOutputOriginalCase = true;
        doc.OptionWriteEmptyNodes = true;
        doc.LoadHtml(originalHtml);

        var warnings = new List<string>();

        // Re-extract using the same traversal to get nodes in order
        var nodesInOrder = new List<(LocalizableString info, HtmlNode node)>();
        int counter = 0;
        int subjectCounter = 0;
        TraverseNodes(doc.DocumentNode, nodesInOrder, ref counter, ref subjectCounter);

        // Build lookup from key to (info, node)
        var nodeLookup = nodesInOrder.ToDictionary(x => x.info.Key, x => x);

        // Build lookup from CSV originals
        var csvLookup = originals.ToDictionary(o => o.Key, o => o);

        foreach (var orig in originals)
        {
            if (!translations.TryGetValue(orig.Key, out var translation))
                continue;

            if (!nodeLookup.TryGetValue(orig.Key, out var match))
            {
                warnings.Add($"Key '{orig.Key}' not found in current HTML — skipped.");
                continue;
            }

            if (match.info.Original != orig.Original)
            {
                warnings.Add($"Key '{orig.Key}': Original text changed. CSV='{orig.Original}', HTML='{match.info.Original}'. Applying translation anyway.");
            }

            var node = match.node;

            if (match.info.Type == "html")
            {
                // Preserve leading/trailing whitespace from original InnerHtml
                var rawHtml = node.InnerHtml;
                var leading = rawHtml[..(rawHtml.Length - rawHtml.TrimStart().Length)];
                var trailing = rawHtml[rawHtml.TrimEnd().Length..];
                node.InnerHtml = leading + translation + trailing;
            }
            else if (match.info.Type == "text")
            {
                // Preserve leading/trailing whitespace from original text node
                var rawText = node.InnerText;
                var trimmed = rawText.Trim();
                var leadIdx = rawText.IndexOf(trimmed[0]);
                var leading = leadIdx > 0 ? rawText[..leadIdx] : "";
                var trailIdx = rawText.LastIndexOf(trimmed[^1]) + 1;
                var trailing = trailIdx < rawText.Length ? rawText[trailIdx..] : "";

                node.InnerHtml = leading + HtmlDocument.HtmlEncode(translation) + trailing;
            }
            else if (match.info.Type == "attribute" && !string.IsNullOrEmpty(match.info.Attribute))
            {
                var attr = node.Attributes[match.info.Attribute];
                if (attr != null)
                {
                    attr.Value = HtmlDocument.HtmlEncode(translation);
                }
                else
                {
                    warnings.Add($"Key '{orig.Key}': Attribute '{match.info.Attribute}' not found on node — skipped.");
                }
            }
        }

        return (doc.DocumentNode.OuterHtml, warnings);
    }

    /// <summary>
    /// Generates a ZIP archive containing localized HTML files for all languages.
    /// </summary>
    public (byte[] zipBytes, Dictionary<string, List<string>> allWarnings) GenerateLocalizedZip(
        string originalHtml,
        string originalFileName,
        byte[] csvBytes)
    {
        var (languages, translations, originals) = ParseTranslatedCsv(csvBytes);
        var allWarnings = new Dictionary<string, List<string>>();

        using var ms = new MemoryStream();
        using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            var baseName = Path.GetFileNameWithoutExtension(originalFileName);
            var extension = Path.GetExtension(originalFileName);

            foreach (var lang in languages)
            {
                var (html, warnings) = GenerateLocalizedHtml(originalHtml, originals, translations[lang]);
                allWarnings[lang] = warnings;

                var entry = archive.CreateEntry($"{baseName}_{lang}{extension}", CompressionLevel.Optimal);
                using var entryStream = entry.Open();
                using var writer = new StreamWriter(entryStream, new UTF8Encoding(false));
                writer.Write(html);
            }
        }

        return (ms.ToArray(), allWarnings);
    }

    private static readonly Regex LocalePattern = new(
        @"^[a-zA-Z]{2,3}(-[a-zA-Z]{2,4})?$", RegexOptions.Compiled);

    /// <summary>
    /// Exports extracted strings to an Excel workbook.
    /// The sheet is named with the source locale. Includes an empty "Translation" column.
    /// </summary>
    public byte[] ExportToXlsx(List<LocalizableString> strings, string locale = "Strings")
    {
        var rows = new List<Dictionary<string, object>>();
        foreach (var s in strings)
        {
            rows.Add(new Dictionary<string, object>
            {
                ["Key"] = s.Key,
                ["Type"] = s.Type,
                ["Original"] = s.Original,
                ["Translation"] = "",
                ["Path"] = s.Path,
                ["Attribute"] = s.Attribute
            });
        }

        using var ms = new MemoryStream();
        var sheets = new Dictionary<string, object> { [locale] = rows };
        MiniExcel.SaveAs(ms, sheets);
        return ms.ToArray();
    }

    /// <summary>
    /// Parses a translated Excel file with a single sheet.
    /// Returns the original LocalizableStrings and a key→translation dictionary.
    /// </summary>
    public (List<LocalizableString> originals, Dictionary<string, string> translations)
        ParseSimpleTranslatedXlsx(byte[] xlsxBytes)
    {
        using var ms = new MemoryStream(xlsxBytes);
        var sheetNames = MiniExcel.GetSheetNames(ms);
        if (sheetNames.Count == 0) return (new(), new());

        ms.Position = 0;
        var rows = MiniExcel.Query(ms, sheetName: sheetNames[0], useHeaderRow: true)
            .Cast<IDictionary<string, object>>().ToList();

        return ParseRowsAsSimpleTranslation(rows);
    }

    /// <summary>
    /// Parses a translated Excel file where each sheet represents a target language.
    /// The sheet name should be a locale code (e.g., "es-ES", "fr-FR").
    /// Each sheet has Key, Type, Original, Translation, Path, Attribute columns.
    /// </summary>
    public (List<string> languages, Dictionary<string, Dictionary<string, string>> translations, List<LocalizableString> originals)
        ParseTranslatedXlsx(byte[] xlsxBytes)
    {
        using var ms = new MemoryStream(xlsxBytes);
        var sheetNames = MiniExcel.GetSheetNames(ms);

        var languages = new List<string>();
        var translations = new Dictionary<string, Dictionary<string, string>>();
        List<LocalizableString>? originals = null;

        foreach (var sheetName in sheetNames)
        {
            ms.Position = 0;
            var rows = MiniExcel.Query(ms, sheetName: sheetName, useHeaderRow: true)
                .Cast<IDictionary<string, object>>().ToList();

            var trimmedName = sheetName.Trim();

            if (trimmedName.Equals("Strings", StringComparison.OrdinalIgnoreCase) ||
                trimmedName.Equals("Template", StringComparison.OrdinalIgnoreCase))
            {
                if (originals == null)
                {
                    var (sheetOriginals, _) = ParseRowsAsSimpleTranslation(rows);
                    originals = sheetOriginals;
                }
                continue;
            }

            if (!LocalePattern.IsMatch(trimmedName))
                continue;

            var (sheetOriginals2, sheetTranslations) = ParseRowsAsSimpleTranslation(rows);

            if (originals == null)
                originals = sheetOriginals2;

            if (sheetTranslations.Count > 0)
            {
                languages.Add(trimmedName);
                translations[trimmedName] = sheetTranslations;
            }
        }

        return (languages, translations, originals ?? new List<LocalizableString>());
    }

    /// <summary>
    /// Parses rows (from MiniExcel) into originals and translations.
    /// Expects columns: Key, Type, Original, Translation, Path, Attribute.
    /// </summary>
    private static (List<LocalizableString> originals, Dictionary<string, string> translations)
        ParseRowsAsSimpleTranslation(List<IDictionary<string, object>> rows)
    {
        var originals = new List<LocalizableString>();
        var translations = new Dictionary<string, string>();

        foreach (var row in rows)
        {
            var key = GetRowString(row, "Key");
            if (string.IsNullOrWhiteSpace(key)) continue;

            originals.Add(new LocalizableString
            {
                Key = key,
                Type = GetRowString(row, "Type"),
                Original = GetRowString(row, "Original"),
                Path = GetRowString(row, "Path"),
                Attribute = GetRowString(row, "Attribute")
            });

            var translation = GetRowString(row, "Translation");
            if (!string.IsNullOrWhiteSpace(translation))
            {
                translations[key] = translation;
            }
        }

        return (originals, translations);
    }

    private static string GetRowString(IDictionary<string, object> row, string columnName)
    {
        if (row.TryGetValue(columnName, out var val) && val != null)
            return val.ToString() ?? string.Empty;
        // Case-insensitive fallback
        var match = row.Keys.FirstOrDefault(k => k.Equals(columnName, StringComparison.OrdinalIgnoreCase));
        if (match != null && row[match] != null)
            return row[match].ToString() ?? string.Empty;
        return string.Empty;
    }

    /// <summary>
    /// Generates a ZIP archive containing localized HTML files from an XLSX file with per-language sheets.
    /// </summary>
    public (byte[] zipBytes, Dictionary<string, List<string>> allWarnings) GenerateLocalizedZipFromXlsx(
        string originalHtml,
        string originalFileName,
        byte[] xlsxBytes)
    {
        var (languages, translations, originals) = ParseTranslatedXlsx(xlsxBytes);
        var allWarnings = new Dictionary<string, List<string>>();

        using var ms = new MemoryStream();
        using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            var baseName = Path.GetFileNameWithoutExtension(originalFileName);
            var extension = Path.GetExtension(originalFileName);

            foreach (var lang in languages)
            {
                var (html, warnings) = GenerateLocalizedHtml(originalHtml, originals, translations[lang]);
                allWarnings[lang] = warnings;

                var entry = archive.CreateEntry($"{baseName}_{lang}{extension}", CompressionLevel.Optimal);
                using var entryStream = entry.Open();
                using var writer = new StreamWriter(entryStream, new UTF8Encoding(false));
                writer.Write(html);
            }
        }

        return (ms.ToArray(), allWarnings);
    }
}
