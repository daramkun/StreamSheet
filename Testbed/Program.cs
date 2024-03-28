// See https://aka.ms/new-console-template for more information

using Hazelnut.StreamSheet;

using var stream = new FileStream("sample.csv", FileMode.Open, FileAccess.Read);
using var reader = new CsvSheetReader(stream, new CsvReaderOptions(headerExists: true));

foreach (var headerColumn in reader.GetHeader() ?? Array.Empty<string>())
{
    Console.Write(headerColumn.PadLeft(16));
}
Console.WriteLine();

// foreach (var record in reader.ReadAsEnumerable())
// {
//     foreach (var column in record)
//         Console.Write(column.PadLeft(16));
//     Console.WriteLine();
// }

reader.ReadAsEnumerable().ToArray();