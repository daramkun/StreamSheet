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

    private string[]? _header;

    private readonly List<string> _record = new();
    
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

    public override string[]? GetHeader() => _header;

    public override string[]? Read()
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
                    if (xmlRowReader is { NodeType: XmlNodeType.EndElement, LocalName: "table-row" })
                        return _record.ToArray();
                    
                    if (xmlRowReader is { NodeType: XmlNodeType.Element, LocalName: "covered-table-cell" })
                        _record.Add(string.Empty);
                    else if (xmlRowReader is { NodeType: XmlNodeType.Element, LocalName: "table-cell" })
                    {
                        var formula = _xmlTableReader.GetAttribute("table:formula");
                        if (formula != null)
                        {
                            _record.Add(formula);
                        }
                        else
                        {
                            var p = _xmlTableReader.ReadSubtree();
                            while (p.Read())
                            {
                                if (p is { NodeType: XmlNodeType.Element, LocalName: "p"})
                                {
                                    _record.Add(p.ReadElementString());
                                    break;
                                }
                            }
                        }
                    }
                }
            }
        }

        if (_record.Count > 0)
            return _record.ToArray();

        return null;
    }

    private static string GetContentAsString(ZipArchiveEntry entry)
    {
        using var stream = entry.Open();
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }
}