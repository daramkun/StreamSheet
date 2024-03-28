namespace Hazelnut.StreamSheet;

public interface ISheetReader : IDisposable
{
    Stream? BaseStream { get; }

    string[]? GetHeader();

    string[]? Read();
    IEnumerable<string[]> ReadAsEnumerable();
}

public interface ITextSheetReader : ISheetReader
{
    TextReader? BaseReader { get; }
}

public abstract class BaseSheetReader : ISheetReader
{
    private bool _isDisposed;
    
    public Stream? BaseStream { get; }
    private readonly bool _leaveOpen;
    
    protected BaseSheetReader(Stream? stream, bool leaveOpen = false)
    {
        BaseStream = stream;
        _leaveOpen = leaveOpen;
    }
    
    ~BaseSheetReader()
    {
        if (_isDisposed)
            return;
        
        Dispose(false);
    }
    
    public void Dispose()
    {
        AssertDisposed();
        
        Dispose(true);
        GC.SuppressFinalize(this);
        
        _isDisposed = true;
    }

    protected virtual void Dispose(bool disposing)
    {
        if (!_leaveOpen)
            BaseStream?.Dispose();
    }

    protected void AssertDisposed()
    {
        if (_isDisposed)
            throw new ObjectDisposedException(GetType().Name);
    }
    
    public abstract string[]? GetHeader();

    public abstract string[]? Read();

    public virtual IEnumerable<string[]> ReadAsEnumerable()
    {
        while (Read() is { } read)
            yield return read;
    }
}

public abstract class BaseTextSheetReader : BaseSheetReader, ITextSheetReader
{
    public TextReader? BaseReader { get; }
    private readonly bool _leaveOpen;
    
    protected BaseTextSheetReader(Stream stream, bool leaveOpen = false)
        : base(stream, leaveOpen)
    {
        BaseReader = new StreamReader(stream, leaveOpen: true);
        _leaveOpen = false;
    }

    protected BaseTextSheetReader(TextReader reader, bool leaveOpen = false)
        : base(reader is StreamReader streamReader ? streamReader.BaseStream : null, true)
    {
        BaseReader = reader;
        _leaveOpen = leaveOpen;
    }

    protected override void Dispose(bool disposing)
    {
        if (!_leaveOpen)
            BaseReader?.Dispose();
        base.Dispose(disposing);
    }
}