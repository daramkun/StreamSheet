namespace Hazelnut.StreamSheet;

public interface ISheetWriter : IDisposable
{
    Stream? BaseStream { get; }

    void Write(IEnumerable<Cell> columns);
}

public interface ITextSheetWriter : IDisposable
{
    TextWriter? BaseWriter { get; }
}

public abstract class BaseSheetWriter : ISheetWriter
{
    private bool _isDisposed;
    
    public Stream? BaseStream { get; }
    private readonly bool _leaveOpen;
    
    protected BaseSheetWriter(Stream? stream, bool leaveOpen = false)
    {
        BaseStream = stream;
        _leaveOpen = leaveOpen;
    }
    
    ~BaseSheetWriter()
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
        {
            BaseStream?.Flush();
            BaseStream?.Dispose();
        }
    }

    protected void AssertDisposed()
    {
        if (_isDisposed)
            throw new ObjectDisposedException(GetType().Name);
    }

    public abstract void Write(IEnumerable<Cell> columns);
}

public abstract class BaseTextSheetWriter : BaseSheetWriter, ITextSheetWriter
{
    public TextWriter? BaseWriter { get; }
    private readonly bool _leaveOpen;
    
    protected BaseTextSheetWriter(Stream stream, bool leaveOpen = false)
        : base(stream, leaveOpen)
    {
        BaseWriter = new StreamWriter(stream, leaveOpen: true);
        _leaveOpen = false;
    }

    protected BaseTextSheetWriter(TextWriter writer, bool leaveOpen = false)
        : base(writer is StreamWriter streamWriter ? streamWriter.BaseStream : null, true)
    {
        BaseWriter = writer;
        _leaveOpen = leaveOpen;
    }

    protected override void Dispose(bool disposing)
    {
        if (!_leaveOpen)
        {
            BaseWriter?.Flush();
            BaseWriter?.Dispose();
        }
        base.Dispose(disposing);
    }
}