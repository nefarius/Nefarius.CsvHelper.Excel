using System;
using System.Globalization;
using System.IO;
using ClosedXML.Excel;
using Xunit;

namespace CsvHelper.Excel.Specs;

public class ReviewRegressionSpecs
{
    [Theory]
    [InlineData(1001, 1000, false)]
    [InlineData(1001, 1000, true)]
    [InlineData(1048576, 16384, false)]
    [InlineData(1048576, 16384, true)]
    public void RejectsOversizedSparseRectangles(int row, int column, bool leaveOpen)
    {
        var stream = new MemoryStream();
        try
        {
            using (var workbook = new XLWorkbook())
            {
                workbook.AddWorksheet("Data").Cell(row, column).Value = "Distant cell";
                workbook.SaveAs(stream);
            }
            stream.Position = 0;
            Assert.Throws<InvalidDataException>(() =>
                new ExcelParser(stream, null, CsvReaderOptions.Default, leaveOpen));
            Assert.Equal(leaveOpen, stream.CanRead);
        }
        finally { stream.Dispose(); }
    }

    [Fact]
    public void KeepsEmptyFieldsForSparseRectanglesWithinTheLimit()
    {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            workbook.AddWorksheet("Data").Cell(1, 16384).Value = "Last column";
            workbook.SaveAs(stream);
        }
        stream.Position = 0;
        using var parser = new ExcelParser(stream, null, new CsvReaderOptions { HasHeader = false });
        Assert.True(parser.Reader.Read());
        Assert.Equal(16384, parser.Reader.FieldCount);
        Assert.Equal("", parser.Reader.GetString(0));
        Assert.Equal("Last column", parser.Reader.GetString(16383));
        Assert.False(parser.Reader.Read());
    }

    [Fact]
    public void PreservesNumberFormatInfo()
    {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            workbook.AddWorksheet("Data").Cell(1, 1).Value = 12.5;
            workbook.SaveAs(stream);
        }
        var provider = (NumberFormatInfo)CultureInfo.InvariantCulture.NumberFormat.Clone();
        provider.NumberDecimalSeparator = ",";
        stream.Position = 0;
        using var parser = new ExcelParser(stream, null, new CsvReaderOptions
        {
            HasHeader = false,
            FormatProvider = provider
        });
        Assert.True(parser.Reader.Read());
        Assert.Equal("12,5", parser.Reader.GetString(0));
        Assert.Equal(12.5m, parser.Reader.Parse<decimal>(0));
    }

    [Fact]
    public void PreservesDateTimeFormatInfo()
    {
        using var stream = new MemoryStream();
        var timestamp = new DateTime(2026, 11, 11, 13, 14, 15);
        using (var workbook = new XLWorkbook())
        {
            workbook.AddWorksheet("Data").Cell(1, 1).Value = timestamp;
            workbook.SaveAs(stream);
        }
        var provider = (DateTimeFormatInfo)CultureInfo.InvariantCulture.DateTimeFormat.Clone();
        provider.ShortDatePattern = "yyyy'Q'MM'Q'dd";
        stream.Position = 0;
        using var parser = new ExcelParser(stream, null, new CsvReaderOptions
        {
            HasHeader = false,
            FormatProvider = provider
        });
        Assert.True(parser.Reader.Read());
        Assert.Equal(timestamp.ToString(provider), parser.Reader.GetString(0));
    }

    [Theory]
    [InlineData(false, "invalid/name", false)]
    [InlineData(true, "invalid/name", false)]
    [InlineData(false, "", false)]
    [InlineData(true, "", false)]
    [InlineData(false, "'name", false)]
    [InlineData(true, "'name", false)]
    [InlineData(false, "12345678901234567890123456789012", false)]
    [InlineData(true, "12345678901234567890123456789012", false)]
    [InlineData(false, "Data", true)]
    [InlineData(true, "Data", true)]
    public void InvalidWriterArgumentsDoNotTouchDestination(bool existingFile, string sheetName, bool invalidOptions)
    {
        var path = Path.Combine(Path.GetTempPath(), $"csvhelper-review-{Guid.NewGuid():N}.xlsx");
        var original = new byte[] { 1, 2, 3, 4 };
        try
        {
            if (existingFile) File.WriteAllBytes(path, original);
            var options = invalidOptions ? new CsvWriterOptions { BufferSize = 255 } : CsvWriterOptions.Default;
            Assert.ThrowsAny<ArgumentException>(() => new ExcelWriter(path, sheetName, options));
            Assert.Equal(existingFile, File.Exists(path));
            if (existingFile) Assert.Equal(original, File.ReadAllBytes(path));
        }
        finally { File.Delete(path); }
    }
}
