namespace Hazelnut.StreamSheet;

public readonly struct OdsSheetReaderOptions(
    bool headerExists = true)
{
    public readonly bool HeaderExists = headerExists;
}