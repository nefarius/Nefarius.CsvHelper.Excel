using System;
using System.Globalization;
using System.IO;
using ClosedXML.Excel;
using CsvHelper;

namespace Nefarius.CsvHelper.Excel;

/// <summary>Opens an Excel worksheet for reading with CsvHelper.</summary>
/// <remarks>Use <see cref="Reader"/> for record mapping and enumeration. The worksheet is buffered as CSV in memory. Dispose this owner after reading.</remarks>
public sealed class ExcelParser : IDisposable
{
    private readonly Stream _source;
    private readonly bool _leaveOpen;
    private readonly MemoryStream _csv = new();
    private bool _disposed;

    /// <summary>Gets the CsvHelper reader over the selected worksheet.</summary>
    public CsvReader Reader { get; }

    /// <summary>Opens the first worksheet of a file.</summary>
    public ExcelParser(string path) : this(path, null, CsvReaderOptions.Default) { }

    /// <summary>Opens a named worksheet of a file.</summary>
    public ExcelParser(string path, string sheetName) : this(path, sheetName, CsvReaderOptions.Default) { }

    /// <summary>Opens a file using the specified culture for cell values and conversions.</summary>
    public ExcelParser(string path, CultureInfo culture) : this(path, null, CsvOptions.Configure(new CsvReaderOptions(), (nameof(CsvReaderOptions.FormatProvider), culture))) { }

    /// <summary>Opens a named worksheet using the specified culture.</summary>
    public ExcelParser(string path, string sheetName, CultureInfo culture)
        : this(path, sheetName, CsvOptions.Configure(new CsvReaderOptions(), (nameof(CsvReaderOptions.FormatProvider), culture))) { }

    /// <summary>Opens a file using CsvHelper reader options.</summary>
    public ExcelParser(string path, string sheetName, CsvReaderOptions options)
        : this(File.OpenRead(path), sheetName, options, false) { }

    /// <summary>Opens the first worksheet of a stream. Set leaveOpen to retain ownership of the stream.</summary>
    public ExcelParser(Stream stream, CultureInfo culture, bool leaveOpen = false)
        : this(stream, null, CsvOptions.Configure(new CsvReaderOptions(), (nameof(CsvReaderOptions.FormatProvider), culture)), leaveOpen) { }

    /// <summary>Opens a named worksheet of a stream using the specified culture.</summary>
    public ExcelParser(Stream stream, string sheetName, CultureInfo culture, bool leaveOpen = false)
        : this(stream, sheetName, CsvOptions.Configure(new CsvReaderOptions(), (nameof(CsvReaderOptions.FormatProvider), culture)), leaveOpen) { }

    /// <summary>Opens a worksheet of a stream using CsvHelper reader options.</summary>
    /// <remarks>CSV dialect and encoding options are normalized for the internal UTF-8 transport.</remarks>
    public ExcelParser(Stream stream, string sheetName, CsvReaderOptions options, bool leaveOpen = false)
    {
        _source = stream ?? throw new ArgumentNullException(nameof(stream));
        _leaveOpen = leaveOpen;
        options ??= CsvReaderOptions.Default;
        try
        {
            using (var workbook = new XLWorkbook(stream))
            {
                var worksheet = string.IsNullOrEmpty(sheetName) ? workbook.Worksheet(1) : workbook.Worksheet(sheetName);
                var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 0;
                var lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 0;
                using var writer = CsvWriter.Create(_csv, CsvWriterOptions.Default, leaveOpen: true);
                for (var row = 1; row <= lastRow; row++)
                {
                    for (var column = 1; column <= lastColumn; column++)
                    {
                        writer.WriteField(worksheet.Cell(row, column).Value.ToString(options.FormatProvider as CultureInfo ?? CultureInfo.InvariantCulture));
                    }
                    writer.EndRecord();
                }
            }
            _csv.Position = 0;
            Reader = CsvReader.Create(_csv, CsvOptions.Configure(options with { },
                (nameof(CsvReaderOptions.Mode), CsvMode.Rfc4180),
                (nameof(CsvReaderOptions.Separator), ','),
                (nameof(CsvReaderOptions.Quote), '"'),
                (nameof(CsvReaderOptions.Escape), '\\'),
                (nameof(CsvReaderOptions.Encoding), System.Text.Encoding.UTF8),
                (nameof(CsvReaderOptions.CommentPrefix), null),
                (nameof(CsvReaderOptions.TrimInsideQuotes), options.TrimInsideQuotes || options.TrimWhitespace == true),
                (nameof(CsvReaderOptions.AllowLineBreaksInQuotes), true)), leaveOpen: true);
        }
        catch
        {
            _csv.Dispose();
            if (!leaveOpen) stream.Dispose();
            throw;
        }
    }

    /// <summary>Disposes the reader and, unless leaveOpen was set, the input stream.</summary>
    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        try { Reader.Dispose(); }
        finally
        {
            _csv.Dispose();
            if (!_leaveOpen) _source.Dispose();
        }
    }
}
