using System.IO.Compression;
using System.Xml;

namespace Hazelnut.StreamSheet;

public class OdsSheetReader : BaseSheetReader
{
    private readonly ZipArchive? _zipArchive;
    private XmlReader? _xmlRootReader;
    private XmlReader? _xmlDocumentContentReader;
    private XmlReader? _xmlBodyReader;
    private XmlReader? _xmlSpreadSheetReader;
    private XmlReader? _xmlTableReader;

    private readonly int _sheetIndex;
    private readonly OdsSheetReaderOptions _options;

    private Cell[]? _header;

    private readonly List<Cell?> _record = [];
    private readonly List<(int, int)> _reservedRows = [];
    private int _rowIndex = 1;
    
    public OdsSheetReader(Stream? stream, int sheetIndex = 0, in OdsSheetReaderOptions options = default, bool leaveOpen = false)
        : base(stream, leaveOpen)
    {
        ArgumentNullException.ThrowIfNull(stream);

        _zipArchive = new ZipArchive(stream, ZipArchiveMode.Read, true);
        if (_zipArchive.GetEntry("mimetype") is not { } mimetype)
            throw new FormatException("This is not Open-Document-Spreadsheet format maybe.");
        if (!GetContentAsString(mimetype).Equals("application/vnd.oasis.opendocument.spreadsheet"))
            throw new FormatException("This is not Open-Document-Spreadsheet format.");

        _sheetIndex = sheetIndex;
        _options = options;

        ReadHeader();
    }

    protected override void Dispose(bool disposing)
    {
        _zipArchive?.Dispose();
        base.Dispose(disposing);
    }

    private void ReadHeader()
    {
        if (!_options.HeaderExists)
            return;

        _header = Read();
    }

    public override Cell[]? GetHeader() => _header;

    public override Cell[]? Read()
    {
        if (_zipArchive == null)
            throw new InvalidOperationException("ZIP Archive is null.");
        
        if (_xmlRootReader == null || _xmlDocumentContentReader == null || _xmlBodyReader == null || _xmlSpreadSheetReader == null || _xmlTableReader == null)
        {
            var entry = _zipArchive.GetEntry("content.xml");
            if (entry == null)
                throw new FormatException("Stream does not contained content.xml");
            
            var stream = entry.Open();
            _xmlRootReader = new XmlTextReader(stream);

            while (_xmlRootReader.Read())
            {
                if (_xmlRootReader is { NodeType: XmlNodeType.Element, LocalName: "document-content" })
                {
                    _xmlDocumentContentReader = _xmlRootReader.ReadSubtree();
                    break;
                }
            }

            if (_xmlDocumentContentReader == null)
                throw new FormatException("content.xml does not contained office:document-content element.");
            
            while (_xmlDocumentContentReader.Read())
            {
                if (_xmlDocumentContentReader is { NodeType: XmlNodeType.Element, LocalName: "body" })
                {
                    _xmlBodyReader = _xmlDocumentContentReader.ReadSubtree();
                    break;
                }
            }

            if (_xmlBodyReader == null)
                throw new FormatException("content.xml does not contained office:body element.");
            
            while (_xmlBodyReader.Read())
            {
                if (_xmlBodyReader is { NodeType: XmlNodeType.Element, LocalName: "spreadsheet" })
                {
                    _xmlSpreadSheetReader = _xmlBodyReader.ReadSubtree();
                    break;
                }
            }

            if (_xmlSpreadSheetReader == null)
                throw new FormatException("content.xml does not contained office:spreadsheet element.");

            var tableIndex = 0;
            while (_xmlSpreadSheetReader.Read())
            {
                if (_xmlSpreadSheetReader is { NodeType: XmlNodeType.Element, LocalName: "table" })
                {
                    if (tableIndex == _sheetIndex)
                    {
                        _xmlTableReader = _xmlBodyReader.ReadSubtree();
                        break;
                    }
                    
                    ++tableIndex;
                }
            }

            if (_xmlTableReader == null)
                throw new FormatException("content.xml does not contained office:table element for _sheetIndex.");
        }

        _record.Clear();

        while (_xmlTableReader.Read())
        {
            if (_xmlTableReader is { NodeType: XmlNodeType.EndElement, LocalName: "table-row" })
                break;
            
            if (_xmlTableReader is { NodeType: XmlNodeType.Element, LocalName: "table-row" })
            {
                var xmlRowReader = _xmlTableReader.ReadSubtree();
                while (xmlRowReader.Read())
                {
                    while (_reservedRows.Contains((_record.Count + 1, _rowIndex)))
                        _record.Add(null);
                    
                    if (xmlRowReader is { NodeType: XmlNodeType.EndElement, LocalName: "table-row" })
                    {
                        ++_rowIndex;
                        return _record.Where(r => r != null).Cast<Cell>().ToArray();
                    }

                    if (xmlRowReader is { NodeType: XmlNodeType.Element, LocalName: "covered-table-cell" })
                        continue;
                    
                    if (xmlRowReader is { NodeType: XmlNodeType.Element, LocalName: "table-cell" })
                    {
                        var columnSpanText = xmlRowReader.GetAttribute("table:number-columns-spanned");
                        var rowSpanText = xmlRowReader.GetAttribute("table:number-rows-spanned");

                        var columnSpan = columnSpanText != null && int.TryParse(columnSpanText, out var cp) ? cp : 1;
                        var rowSpan = rowSpanText != null && int.TryParse(rowSpanText, out var rp) ? rp : 1;
                        
                        var formula = _xmlTableReader.GetAttribute("table:formula");
                        if (formula != null)
                        {
                            _record.Add(new Cell(_record.Count + 1, _rowIndex, formula, columnSpan, rowSpan));
                        }
                        else
                        {
                            var found = false;
                            var p = _xmlTableReader.ReadSubtree();
                            while (p.Read())
                            {
                                if (p is { NodeType: XmlNodeType.Element, LocalName: "p"})
                                {
                                    _record.Add(new Cell(_record.Count + 1, _rowIndex, p.ReadElementString(), columnSpan, rowSpan));
                                    found = true;
                                    break;
                                }
                            }
                            
                            if (!found)
                                _record.Add(new Cell(_record.Count + 1, _rowIndex, string.Empty, columnSpan, rowSpan));
                        }

                        for (var y = 0; y < rowSpan; ++y)
                        {
                            for (var x = 1; x < columnSpan; ++x)
                            {
                                if (y == 0)
                                    _record.Add(null);
                                else
                                    _reservedRows.Add((_record.Count, _rowIndex + y + 1));
                            }
                        }
                    }
                }
            }
        }

        if (_record.Count > 0)
            return _record.Where(r => r != null).Cast<Cell>().ToArray();

        return null;
    }

    private static string GetContentAsString(ZipArchiveEntry entry)
    {
        using var stream = entry.Open();
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }
}