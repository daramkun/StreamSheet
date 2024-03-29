using System.Diagnostics.CodeAnalysis;
using System.Text;

namespace Hazelnut.StreamSheet;

[RequiresDynamicCode("RFC4180 Comma-separated value Format Reader")]
public sealed class CsvSheetReader : BaseTextSheetReader
{
    private readonly CsvSheetReaderOptions _options;
    private Cell[]? _header;
    private int _columnCount = -1;
    private int _rowCount = 1;

    private readonly StringBuilder _builder = new();
    private readonly List<Cell> _record = [];

    public CsvSheetReaderOptions Options => _options;

    public CsvSheetReader(Stream stream, in CsvSheetReaderOptions options = default, bool leaveOpen = false)
        : base(stream, leaveOpen)
    {
        _options = options;
        ReadHeader();
    }

    public CsvSheetReader(TextReader reader, in CsvSheetReaderOptions options = default, bool leaveOpen = false)
        : base(reader, leaveOpen)
    {
        _options = options;
        ReadHeader();
    }

    private void ReadHeader()
    {
        if (!_options.HeaderExists)
            return;

        _header = Read();
    }

    public override Cell[]? GetHeader()
    {
        AssertDisposed();
        return _header;
    }

    public override Cell[]? Read()
    {
        AssertDisposed();
        
        var reader = BaseReader;
        if (reader == null)
            return null;

        _builder.Clear();
        _record.Clear();

        var state = CsvReaderState.None;

        int ch;
        while ((ch = reader.Read()) != -1)
        {
            if (ch == '\r')
                continue;
            
            switch (state)
            {
                case CsvReaderState.None:
                    if (ch == _options.ColumnDelimiter)
                        state = CsvReaderState.ColumnEnd;
                    else if (ch == _options.RecordDelimiter)
                        state = CsvReaderState.RecordEnd;
                    else if (_options.QuoteBindable && ch == '"')
                        state = CsvReaderState.QuoteBoundField;
                    else
                    {
                        state = CsvReaderState.UnboundField;
                        _builder.Append((char)ch);
                    }
                    
                    break;
                
                case CsvReaderState.UnboundField:
                    if (ch == _options.ColumnDelimiter)
                        state = CsvReaderState.ColumnEnd;
                    else if (ch == _options.RecordDelimiter)
                        state = CsvReaderState.RecordEnd;
                    else
                        _builder.Append((char)ch);
                    break;
                
                case CsvReaderState.QuoteBoundField:
                    if ((ch == _options.RecordDelimiter || ch == _options.ColumnDelimiter) &&
                        _builder[^1] == '"')
                    {
                        _builder.Remove(_builder.Length - 1, 1);
                        if (ch == _options.ColumnDelimiter)
                            state = CsvReaderState.ColumnEnd;
                        else if (ch == _options.RecordDelimiter)
                            state = CsvReaderState.RecordEnd;
                        else
                            throw new InvalidOperationException("This logic's ch variable must have record delimiter or column delimiter.");
                    }
                    else
                    {
                        _builder.Append((char)ch);
                        if (ch != '"' && IsCharExists(_builder, ^2, true, '"'))
                        {
                            var quoteCount = 1;
                            while (IsCharExists(_builder, ^(quoteCount + 2), _options.BackslashDoubleQuotesForBound, '\\') ||
                                   IsCharExists(_builder, ^(quoteCount + 2), _options.DoubleDoubleQuotesForBound, '"'))
                            {
                                var backslash = IsCharExists(_builder, ^(quoteCount + 2), _options.BackslashDoubleQuotesForBound, '\\');
                                if (backslash && IsCharExists(_builder, ^(quoteCount + 3), _options.BackslashDoubleQuotesForBound, '\\'))
                                    break;
                                ++quoteCount;
                            }

                            if (quoteCount % 2 != 0)
                                throw new FormatException("Double-quotes with no double double-quotes or backslash.");
                        }
                    }

                    break;
            }

            if (state is CsvReaderState.ColumnEnd or CsvReaderState.RecordEnd)
            {
                if (_builder.Length > 0)
                {
                    if (_options.Trimming)
                    {
                        var i = 0;
                        while (_builder[i] == ' ' || _builder[i] == '\t' || _builder[i] == '\0')
                            ++i;
                        if (i > 0)
                            _builder.Remove(0, i);
                    
                        i = 0;
                        while (_builder[_builder.Length - i - 1] == ' ' || _builder[_builder.Length - i - 1] == '\t' ||
                               _builder[_builder.Length - i - 1] == '\0')
                            ++i;
                        if (i > 0)
                            _builder.Remove(_builder.Length - i - 1, i);
                    }
                }

                var cellText = _builder.Length > 0
                    ? _builder.ToString()
                    : string.Empty;
                var cell = new Cell(_record.Count + 1, _rowCount, cellText);
                _record.Add(cell);

                _builder.Clear();

                if (state == CsvReaderState.RecordEnd)
                    break;

                state = CsvReaderState.None;
            }
        }

        if (_record.Count == 0 || (_record.Count == 1 && string.IsNullOrEmpty(_record[0].Value as string)))
            return null;

        if (_options.SameColumnCount && _columnCount != -1)
        {
            if (_record.Count != _columnCount)
                throw new FormatException($"Column count mismatched: {_columnCount} / {_record.Count}");
        }

        _columnCount = _record.Count;
        ++_rowCount;
        
        return _record.ToArray();
    }

    private enum CsvReaderState
    {
        None,
        UnboundField,
        QuoteBoundField,
        ColumnEnd,
        RecordEnd,
    }

    private static bool IsCharExists(StringBuilder builder, Index index, bool condition, char ch)
    {
        if (!condition)
            return false;

        if (builder.Length <= index.Value)
            return false;

        return builder[index] == ch;
    }
}