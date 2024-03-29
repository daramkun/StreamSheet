using System.Text;

namespace Hazelnut.StreamSheet;

public class CsvSheetWriter : BaseTextSheetWriter
{
    private readonly CsvWriterOptions _options;
    private int _columnCount = -1;

    private readonly StringBuilder _builder = new();

    public CsvWriterOptions Options => _options;
    
    public CsvSheetWriter(Stream stream, in CsvWriterOptions options = default, bool leaveOpen = false)
        : base(stream, leaveOpen)
    {
    }

    public CsvSheetWriter(TextWriter writer, in CsvWriterOptions options = default, bool leaveOpen = false)
        : base(writer, leaveOpen)
    {
    }

    public override void Write(IEnumerable<string> columns)
    {
        if (BaseWriter == null)
            return;
        
        _builder.Clear();

        var count = 0;

        foreach (var column in columns)
        {
            if (count != 0)
                _builder.Append(_options.ColumnDelimiter);

            if (!string.IsNullOrEmpty(column))
            {
                if (!_options.QuoteBindable && column.Contains(_options.ColumnDelimiter))
                    throw new ArgumentOutOfRangeException(nameof(columns), "Column value contains column delimiter but Quote for value binding is not disabled.");
                if (!_options.QuoteBindable && column.Contains(_options.RecordDelimiter))
                    throw new ArgumentOutOfRangeException(nameof(columns), "Column value contains record delimiter but Quote for value binding is not disabled.");

                if (_options.QuoteBindable &&
                    (column.Contains(_options.ColumnDelimiter) || column.Contains(_options.RecordDelimiter) || column.Contains('"')))
                {
                    _builder.Append('"');
                    foreach (var ch in column)
                    {
                        if (ch == '"')
                        {
                            if (_options.DoubleDoubleQuotesForBound)
                                _builder.Append('"').Append('"');
                            else if (_options.BackslashDoubleQuotesForBound)
                                _builder.Append('\\').Append('"');
                        }
                        else
                            _builder.Append(ch);
                    }
                    _builder.Append('"');
                }
                else
                {
                    _builder.Append(column);
                }
            }
            
            ++count;
        }

        if (_columnCount != -1 && _columnCount != count && _options.SameColumnCount)
            throw new ArgumentOutOfRangeException(nameof(columns), "Column count is mismatch.");

        _columnCount = count;

        if (_builder.Length == 0)
            return;

        if (_options.RecordDelimiter == '\n')
            _builder.Append(CRLF);
        else
            _builder.Append(_options.RecordDelimiter);
        
        BaseWriter.Write(_builder);
    }

    private static readonly string CRLF = "\r\n";
}