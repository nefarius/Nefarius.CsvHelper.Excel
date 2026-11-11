using System;
using System.Globalization;
using System.IO;
using System.Threading.Tasks;
using ClosedXML.Excel;
using CsvHelper;

namespace Nefarius.CsvHelper.Excel;

/// <summary>Writes CsvHelper records into an Excel workbook when disposed.</summary>
/// <remarks>Use <see cref="Writer"/> to write records. Intermediate CSV is buffered in memory until disposal.</remarks>
public sealed class ExcelWriter : IDisposable, IAsyncDisposable
{
    private readonly Stream _destination;
    private readonly bool _leaveOpen;
    private readonly string _sheetName;
    private readonly MemoryStream _csv = new();
    private bool _disposed;

    /// <summary>Gets the CsvHelper writer used to format fields and records.</summary>
    public CsvWriter Writer { get; }

    /// <summary>Creates or overwrites a file with an export worksheet.</summary>
    public ExcelWriter(string path) : this(path, "export", CsvWriterOptions.Default) { }

    /// <summary>Creates or overwrites a file using the specified culture.</summary>
    public ExcelWriter(string path, CultureInfo culture) : this(path, "export", CsvOptions.Configure(new CsvWriterOptions(), (nameof(CsvWriterOptions.FormatProvider), culture))) { }

    /// <summary>Creates or overwrites a file with a named worksheet.</summary>
    public ExcelWriter(string path, string sheetName) : this(path, sheetName, CsvWriterOptions.Default) { }

    /// <summary>Creates or overwrites a file with a named worksheet and culture.</summary>
    public ExcelWriter(string path, string sheetName, CultureInfo culture)
        : this(path, sheetName, CsvOptions.Configure(new CsvWriterOptions(), (nameof(CsvWriterOptions.FormatProvider), culture))) { }

    /// <summary>Creates or overwrites a file using CsvHelper writer options.</summary>
    public ExcelWriter(string path, string sheetName, CsvWriterOptions options)
        : this(File.Create(path), sheetName, options, false) { }

    /// <summary>Writes to a stream using the specified culture. Set leaveOpen to retain stream ownership.</summary>
    public ExcelWriter(Stream stream, CultureInfo culture, bool leaveOpen = false)
        : this(stream, "export", CsvOptions.Configure(new CsvWriterOptions(), (nameof(CsvWriterOptions.FormatProvider), culture)), leaveOpen) { }

    /// <summary>Writes a named worksheet to a stream using the specified culture.</summary>
    public ExcelWriter(Stream stream, string sheetName, CultureInfo culture, bool leaveOpen = false)
        : this(stream, sheetName, CsvOptions.Configure(new CsvWriterOptions(), (nameof(CsvWriterOptions.FormatProvider), culture)), leaveOpen) { }

    /// <summary>Writes a named worksheet to a stream using CsvHelper writer options.</summary>
    /// <remarks>CSV dialect options are normalized for the internal transport; formatting and mapping options apply.</remarks>
    public ExcelWriter(Stream stream, string sheetName, CsvWriterOptions options, bool leaveOpen = false)
    {
        _destination = stream ?? throw new ArgumentNullException(nameof(stream));
        _sheetName = sheetName;
        _leaveOpen = leaveOpen;
        try
        {
            // Validate the worksheet name before accepting records.
            using var workbook = new XLWorkbook();
            workbook.AddWorksheet(sheetName);
            Writer = CsvWriter.Create(_csv, CsvOptions.Configure((options ?? CsvWriterOptions.Default) with { },
                (nameof(CsvWriterOptions.Mode), CsvMode.Rfc4180),
                (nameof(CsvWriterOptions.Separator), ','),
                (nameof(CsvWriterOptions.Escape), '\\'),
                (nameof(CsvWriterOptions.NewLine), "\r\n"),
                (nameof(CsvWriterOptions.CommentPrefix), null)), leaveOpen: true);
        }
        catch
        {
            _csv.Dispose();
            if (!leaveOpen) stream.Dispose();
            throw;
        }
    }

    private void SaveWorkbook()
    {
        _csv.Position = 0;
        using var reader = CsvReader.Create(_csv, CsvOptions.Configure(new CsvReaderOptions(),
            (nameof(CsvReaderOptions.HasHeader), false),
            (nameof(CsvReaderOptions.Separator), ','),
            (nameof(CsvReaderOptions.SkipEmptyLines), false),
            (nameof(CsvReaderOptions.TrimWhitespace), false)), leaveOpen: true);
        using var workbook = new XLWorkbook();
        var worksheet = workbook.AddWorksheet(_sheetName);
        var row = 1;
        while (reader.Read())
        {
            for (var column = 0; column < reader.FieldCount; column++)
            {
                worksheet.Cell(row, column + 1).Value = reader.GetString(column);
            }
            row++;
        }
        workbook.SaveAs(_destination);
    }

    /// <summary>Finishes the records, saves the workbook and releases owned resources.</summary>
    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        try
        {
            Writer.Dispose();
            SaveWorkbook();
            _destination.Flush();
        }
        finally
        {
            _csv.Dispose();
            if (!_leaveOpen) _destination.Dispose();
        }
    }

    /// <summary>Flushes the records asynchronously, saves the workbook and releases owned resources.</summary>
    /// <remarks>ClosedXML workbook serialization is synchronous.</remarks>
    public async ValueTask DisposeAsync()
    {
        if (_disposed) return;
        _disposed = true;
        try
        {
            await Writer.DisposeAsync().ConfigureAwait(false);
            SaveWorkbook();
            await _destination.FlushAsync().ConfigureAwait(false);
        }
        finally
        {
            _csv.Dispose();
            if (!_leaveOpen) _destination.Dispose();
        }
    }
}
