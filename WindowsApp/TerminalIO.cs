using System.Collections.Concurrent;

namespace USBDriveInventoryCollector;

internal interface ITerminalIO
{
    bool KeyAvailable { get; }
    char ReadKey();
    string? ReadLine();
    void Write(string value);
    void WriteLine(string value);
    void Clear();
}

internal sealed class ConsoleTerminalIO : ITerminalIO
{
    public bool KeyAvailable => Console.KeyAvailable;
    public char ReadKey() => Console.ReadKey(true).KeyChar;
    public string? ReadLine() => Console.ReadLine();
    public void Write(string value) => Console.Write(value);
    public void WriteLine(string value) => Console.WriteLine(value);
    public void Clear() => Console.Clear();
}

internal sealed class EmbeddedTerminalIO : ITerminalIO, IDisposable
{
    private readonly BlockingCollection<string> _lines = new();
    private int _waitingForLine;
    public event Action<string>? Output;
    public event Action? Cleared;
    public bool KeyAvailable => _lines.Count > 0;
    public bool WaitingForLine => Volatile.Read(ref _waitingForLine) != 0;
    public char ReadKey() => ReadLine()?.Trim().FirstOrDefault() ?? '\0';
    public string? ReadLine()
    {
        Interlocked.Exchange(ref _waitingForLine, 1);
        try { return _lines.Take(); }
        catch (InvalidOperationException) { return null; }
        finally { Interlocked.Exchange(ref _waitingForLine, 0); }
    }
    public void Submit(string value)
    {
        if (!_lines.IsAddingCompleted) try { _lines.Add(value); } catch (InvalidOperationException) { }
    }
    public void Write(string value) => Output?.Invoke(value);
    public void WriteLine(string value) => Output?.Invoke(value + Environment.NewLine);
    public void Clear() => Cleared?.Invoke();
    public void Close() => _lines.CompleteAdding();
    public void Dispose() { Close(); _lines.Dispose(); }
}
