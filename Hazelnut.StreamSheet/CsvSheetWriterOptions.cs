namespace Hazelnut.StreamSheet;

public struct CsvSheetWriterOptions(
    char columnDelimiter = ',',
    char recordDelimiter = '\n',
    bool quoteBindable = true,
    bool doubleDoubleQuotesForBound = true,
    bool backslashDoubleQuotesForBound = false,
    bool sameColumnCount = true)
{
    public readonly char ColumnDelimiter = columnDelimiter;
    public readonly char RecordDelimiter = recordDelimiter;
    public readonly bool QuoteBindable = quoteBindable;
    public readonly bool DoubleDoubleQuotesForBound = doubleDoubleQuotesForBound;
    public readonly bool BackslashDoubleQuotesForBound = backslashDoubleQuotesForBound;
    
    public readonly bool SameColumnCount = sameColumnCount;
}