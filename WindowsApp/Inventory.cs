using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

namespace USBDriveInventoryCollector;

internal sealed class DriveRecord
{
    public Dictionary<string, string> Values { get; } = new(StringComparer.OrdinalIgnoreCase);
    public string this[string key]
    {
        get => Values.TryGetValue(key, out var value) && !string.IsNullOrWhiteSpace(value) ? value : "N/A";
        set => Values[key] = string.IsNullOrWhiteSpace(value) ? "N/A" : value.Trim();
    }
    public DriveRecord Copy()
    {
        var copy = new DriveRecord();
        foreach (var (key, value) in Values) copy[key] = value;
        return copy;
    }
}

internal sealed class InventoryBook
{
    public static readonly (string Key, string Header)[] Catalog = [
        ("Manufacturer", "Manufacturer"), ("Model", "Model"), ("SerialNumber", "Serial Number"),
        ("Capacity", "Reported Capacity"), ("Type", "Type"),
        ("Interface", "Interface"), ("FirmwareVersion", "Firmware Version"),
        ("ModelFamily", "Model Family"), ("FormFactor", "Form Factor"),
        ("RotationRate", "Rotation Rate (RPM)"), ("CapacityBytes", "Capacity (Bytes)"),
        ("LogicalBlockSize", "Logical Sector Size (Bytes)"),
        ("PhysicalBlockSize", "Physical Sector Size (Bytes)"),
        ("AtaVersion", "ATA Version"), ("SataVersion", "SATA Version"),
        ("Protocol", "Reported Protocol"), ("Transport", "Probe Transport")
    ];
    public static readonly string[] Core = ["Manufacturer", "Model", "SerialNumber", "Capacity", "Type"];
    public List<string> Columns { get; private set; } = [..Core];
    public List<DriveRecord> Records { get; } = [];
    public string Path { get; }
    private static readonly XNamespace X = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace DocRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    public InventoryBook(string path) { Path = path; }
    public bool HasSerial(string serial, int? exceptIndex = null) => serial != "N/A" && Records.Where((_, i) => i != exceptIndex).Any(r => string.Equals(r["SerialNumber"], serial, StringComparison.OrdinalIgnoreCase));
    public void OpenOrCreate()
    {
        Directory.CreateDirectory(System.IO.Path.GetDirectoryName(Path)!);
        if (!File.Exists(Path)) { Save(); return; }
        bool legacyHeader;
        {
            using var archive = ZipFile.OpenRead(Path);
            var sheet = archive.GetEntry("xl/worksheets/sheet1.xml") ?? throw new InvalidDataException("Workbook has no first worksheet.");
            var shared = new List<string>();
            if (archive.GetEntry("xl/sharedStrings.xml") is { } strings)
            {
                using var stream = strings.Open();
                shared = XDocument.Load(stream).Descendants(X + "si")
                    .Select(si => string.Concat(si.Descendants(X + "t").Select(t => t.Value))).ToList();
            }
            using var sheetStream = sheet.Open();
            var doc = XDocument.Load(sheetStream);
            var rows = doc.Descendants(X + "sheetData").Elements(X + "row").ToList();
            var header = rows.FirstOrDefault(r => (string?)r.Attribute("r") == "1") ?? throw new InvalidDataException("Workbook has no header row.");
            var mapping = new Dictionary<string, string>();
            var loaded = new List<string>();
            legacyHeader = false;
            foreach (var cell in header.Elements(X + "c"))
            {
                var label = CellText(cell, shared).Trim();
                if (label.Length == 0) continue;
                // Older inventories called Manufacturer "Make". Read either header into
                // the same field, then write the new header when the book is saved.
                var match = Catalog.FirstOrDefault(c => c.Header == label || (label == "Make" && c.Key == "Manufacturer") || (label == "Capacity" && c.Key == "Capacity"));
                if (match.Key is null || loaded.Contains(match.Key)) throw new InvalidDataException($"Unsupported or duplicate column '{label}'.");
                var reference = (string?)cell.Attribute("r") ?? "";
                if (!System.Text.RegularExpressions.Regex.IsMatch(reference, "^[A-Z]+1$")) throw new InvalidDataException("Invalid header cell reference.");
                mapping[reference[..^1]] = match.Key;
                loaded.Add(match.Key);
                legacyHeader |= label == "Make";
            }
            foreach (var key in Core) if (!loaded.Contains(key)) throw new InvalidDataException($"Required column '{Header(key)}' is missing.");
            foreach (var row in rows.Skip(1))
            {
                var number = (string?)row.Attribute("r") ?? "";
                if (!int.TryParse(number, out var rowNo) || rowNo <= 1) continue;
                var record = new DriveRecord();
                bool populated = false;
                foreach (var key in loaded) record[key] = "N/A";
                foreach (var cell in row.Elements(X + "c"))
                {
                    var reference = (string?)cell.Attribute("r") ?? "";
                    var column = System.Text.RegularExpressions.Regex.Match(reference, "^[A-Z]+(?=\\d+$)").Value;
                    var value = CellText(cell, shared);
                    if (mapping.TryGetValue(column, out var key)) { record[key] = value; populated |= value.Length > 0; }
                    else if (!string.IsNullOrWhiteSpace(value)) throw new InvalidDataException($"Data without a column header in row {rowNo}.");
                }
                if (populated) Records.Add(record);
            }
            Columns = loaded;
        }
        if (legacyHeader) Save();
    }
    public static string Header(string key) => Catalog.First(c => c.Key == key).Header;
    private static string CellText(XElement cell, List<string> shared)
    {
        var type = (string?)cell.Attribute("t");
        if (type == "inlineStr") return string.Concat(cell.Element(X + "is")?.Descendants(X + "t").Select(x => x.Value) ?? []);
        var value = (string?)cell.Element(X + "v") ?? "";
        return type == "s" && int.TryParse(value, out var n) && n >= 0 && n < shared.Count ? shared[n] : value;
    }
    public int Add(DriveRecord record)
    {
        if (HasSerial(record["SerialNumber"])) throw new InvalidOperationException("This serial number is already in the workbook.");
        Records.Add(record);
        try { Save(); }
        catch { Records.RemoveAt(Records.Count - 1); throw; }
        return Records.Count + 1;
    }
    public void Update(int index, DriveRecord record)
    {
        if (index < 0 || index >= Records.Count) throw new ArgumentOutOfRangeException(nameof(index));
        if (HasSerial(record["SerialNumber"], index)) throw new InvalidOperationException("This serial number is already in the workbook.");
        var previous = Records[index];
        Records[index] = record;
        try { Save(); }
        catch { Records[index] = previous; throw; }
    }
    public InventoryBook AtLocation(string destination)
    {
        destination = System.IO.Path.GetFullPath(destination);
        if (string.Equals(destination, System.IO.Path.GetFullPath(Path), StringComparison.OrdinalIgnoreCase)) return this;
        Directory.CreateDirectory(System.IO.Path.GetDirectoryName(destination)!);
        // Copy the current workbook when choosing a new path; never overwrite an existing workbook.
        if (!File.Exists(destination)) File.Copy(Path, destination);
        var book = new InventoryBook(destination);
        book.OpenOrCreate();
        return book;
    }
    public void ChangeColumns(List<string> selected)
    {
        if (!Core.SequenceEqual(selected.Take(Core.Length)) || selected.Distinct().Count() != selected.Count ||
            selected.Except(Catalog.Select(c => c.Key)).Any()) throw new InvalidOperationException("Invalid workbook columns.");
        var backup = Path + ".before-setup-" + Guid.NewGuid().ToString("N") + ".xlsx";
        File.Copy(Path, backup);
        var previous = Columns;
        Columns = selected;
        try { Save(); }
        catch { Columns = previous; throw; }
    }
    public void Save()
    {
        var sheet = new XElement(X + "worksheet",
            new XElement(X + "sheetViews", new XElement(X + "sheetView", new XAttribute("workbookViewId", 0),
                new XElement(X + "pane", new XAttribute("ySplit", 1), new XAttribute("topLeftCell", "A2"), new XAttribute("activePane", "bottomLeft"), new XAttribute("state", "frozen")))),
            new XElement(X + "sheetFormatPr", new XAttribute("defaultRowHeight", 15)),
            new XElement(X + "cols", Columns.Select((key, i) => new XElement(X + "col", new XAttribute("min", i + 1), new XAttribute("max", i + 1), new XAttribute("width", key == "Model" ? 35 : 25), new XAttribute("customWidth", 1)))),
            new XElement(X + "sheetData", new[] { Row(1, Columns.Select(Header).ToArray(), true) }
                .Concat(Records.Select((r, i) => Row(i + 2, Columns.Select(key => r[key]).ToArray(), false)))),
            Records.Count > 0 ? new XElement(X + "autoFilter", new XAttribute("ref", $"A1:{ColumnName(Columns.Count)}{Records.Count + 1}")) : null);
        var entries = new Dictionary<string, string> {
            ["xl/workbook.xml"] = new XDocument(new XElement(X + "workbook", new XAttribute(XNamespace.Xmlns + "r", DocRel),
                new XElement(X + "bookViews", new XElement(X + "workbookView")),
                new XElement(X + "sheets", new XElement(X + "sheet", new XAttribute("name", "Inventory"), new XAttribute("sheetId", 1), new XAttribute(DocRel + "id", "rId1"))))).ToString(),
            ["xl/styles.xml"] = """<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>""",
            ["xl/worksheets/sheet1.xml"] = new XDocument(sheet).ToString()
        };
        Exception? last = null;
        for (int attempt = 1; attempt <= 5; attempt++)
        {
            var temp = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Path)!, $".{System.IO.Path.GetFileNameWithoutExtension(Path)}.{Guid.NewGuid():N}.xlsx");
            try
            {
                using (var document = SpreadsheetDocument.Create(temp, SpreadsheetDocumentType.Workbook))
                {
                    var workbook = document.AddWorkbookPart();
                    using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(entries["xl/workbook.xml"]))) workbook.FeedData(stream);
                    var worksheet = workbook.AddNewPart<WorksheetPart>("rId1");
                    using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(entries["xl/worksheets/sheet1.xml"]))) worksheet.FeedData(stream);
                    var styles = workbook.AddNewPart<WorkbookStylesPart>("rId2");
                    using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(entries["xl/styles.xml"]))) styles.FeedData(stream);
                }
                using (var check = SpreadsheetDocument.Open(temp, false))
                {
                    var errors = new OpenXmlValidator(FileFormatVersions.Office2007).Validate(check).Take(5)
                        .Select(error => $"{error.Path?.XPath}: {error.Description}").ToArray();
                    if (errors.Length > 0) throw new InvalidDataException("Workbook validation failed: " + string.Join("; ", errors));
                }
                if (File.Exists(Path)) { try { File.Replace(temp, Path, null); } catch (PlatformNotSupportedException) { File.Move(temp, Path, true); } }
                else File.Move(temp, Path);
                return;
            }
            catch (Exception ex) { last = ex; Thread.Sleep(750 * attempt); }
            finally { if (File.Exists(temp)) File.Delete(temp); }
        }
        throw new IOException("Workbook could not be saved after five attempts.", last);
    }
    private static XElement Row(int number, string[] values, bool header) => new(X + "row", new XAttribute("r", number), values.Select((value, i) =>
        new XElement(X + "c", new XAttribute("r", ColumnName(i + 1) + number), new XAttribute("t", "inlineStr"), header ? new XAttribute("s", 1) : null,
            new XElement(X + "is", new XElement(X + "t", value)))));
    private static string ColumnName(int index) { string s = ""; while (index > 0) { index--; s = (char)('A' + index % 26) + s; index /= 26; } return s; }
}
