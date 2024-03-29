// See https://aka.ms/new-console-template for more information

using Hazelnut.StreamSheet;

// using var stream = new FileStream("sample.csv", FileMode.Open, FileAccess.Read);
// using var reader = new CsvSheetReader(stream, new CsvSheetReaderOptions(headerExists: true));
using var stream = new FileStream("sample.ods", FileMode.Open, FileAccess.Read);
using var reader = new OdsSheetReader(stream, 0, new OdsSheetReaderOptions(headerExists: true));

foreach (var headerColumn in reader.GetHeader() ?? Array.Empty<Cell>())
{
    Console.Write("{0}({1})", headerColumn.ToString(), headerColumn.ToAddress().PadRight(8));
}
Console.WriteLine();

foreach (var record in reader.ReadAsEnumerable())
{
    foreach (var column in record)
        Console.Write("{0}({1})", column.ToString(), column.ToAddress().PadRight(8));
    Console.WriteLine();
}