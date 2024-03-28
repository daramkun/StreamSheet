namespace Hazelnut.StreamSheet;

public readonly struct CsvReaderOptions(
    char columnDelimiter = ',',
    char recordDelimiter = '\n',
    bool trimming = false,
    bool quoteBindable = true,
    bool doubleDoubleQuotesForBound = true,
    bool backslashDoubleQuotesForBound = false,
    bool headerExists = true,
    bool sameColumnCount = true)
{
    public readonly char ColumnDelimiter = columnDelimiter;
    public readonly char RecordDelimiter = recordDelimiter;
    public readonly bool Trimming = trimming;
    public readonly bool QuoteBindable = quoteBindable;
    public readonly bool DoubleDoubleQuotesForBound = doubleDoubleQuotesForBound;
    public readonly bool BackslashDoubleQuotesForBound = backslashDoubleQuotesForBound;
    
    public readonly bool HeaderExists = headerExists;
    public readonly bool SameColumnCount = sameColumnCount;
}