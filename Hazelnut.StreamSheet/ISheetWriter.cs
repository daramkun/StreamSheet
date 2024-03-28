namespace Hazelnut.StreamSheet;

public interface ISheetWriter : IDisposable
{
    Stream? BaseStream { get; }

    void Write(IEnumerable<string> columns);
}

public interface ITextSheetWriter : IDisposable
{
    TextWriter? BaseWriter { get; }
}