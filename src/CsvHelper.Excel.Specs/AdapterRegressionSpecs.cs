using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ClosedXML.Excel;
using CsvHelper.Excel.Specs.Common;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace CsvHelper.Excel.Specs;

public class AdapterRegressionSpecs
{
    [Theory]
    [InlineData("Comma, quote \" and\r\nnew line")]
    [InlineData("  leading and trailing whitespace  ")]
    [InlineData("Grüße 日本語 😀")]
    [InlineData("=SUM(A1:A2)")]
    public void RoundTripsCellText(string text)
    {
        using var stream = new MemoryStream();
        using (var owner = new ExcelWriter(stream, CultureInfo.InvariantCulture, leaveOpen: true))
        {
            owner.Writer.WriteRecords(new[] { new Person { Name = text, Age = null } });
        }
        stream.Position = 0;
        using var parser = new ExcelParser(stream, CultureInfo.InvariantCulture, leaveOpen: true);
        var record = Assert.Single(parser.Reader.GetRecords<Person>());
        // ClosedXML normalizes Excel cell line endings to LF.
        Assert.Equal(text.Replace("\r\n", "\n"), record.Name);
        Assert.Null(record.Age);
    }

    [Fact]
    public async Task AsyncDisposalSavesWorkbookAndLeavesStreamOpen()
    {
        using var stream = new MemoryStream();
        var owner = new ExcelWriter(stream, CultureInfo.InvariantCulture, leaveOpen: true);
        await owner.Writer.WriteRecordsAsync(new[] { new Person { Name = "Async", Age = 42 } });
        await owner.DisposeAsync();
        await owner.DisposeAsync();
        Assert.True(stream.CanRead);
        stream.Position = 0;
        using var parser = new ExcelParser(stream, CultureInfo.InvariantCulture, leaveOpen: true);
        var count = 0;
        await foreach (var record in parser.Reader.GetRecordsAsync<Person>())
        {
            count++;
            Assert.Equal("Async", record.Name);
            Assert.Equal(42, record.Age);
        }
        Assert.Equal(1, count);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RespectsStreamOwnership(bool leaveOpen)
    {
        var stream = new MemoryStream();
        using (var writer = new ExcelWriter(stream, CultureInfo.InvariantCulture, leaveOpen))
        {
            writer.Writer.WriteRecords(new[] { new Person { Name = "Owner", Age = 1 } });
        }
        Assert.Equal(leaveOpen, stream.CanRead);
        if (leaveOpen)
        {
            stream.Position = 0;
            using (var parser = new ExcelParser(stream, CultureInfo.InvariantCulture, leaveOpen))
            {
                Assert.Single(parser.Reader.GetRecords<Person>());
            }
            Assert.True(stream.CanRead);
            stream.Position = 0;
            using (var parser = new ExcelParser(stream, CultureInfo.InvariantCulture))
            {
                Assert.Single(parser.Reader.GetRecords<Person>());
            }
            Assert.False(stream.CanRead);
        }
        stream.Dispose();
    }

    [Fact]
    public void RetainsLeadingEmptyColumnsAndLastUsedColumn()
    {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.AddWorksheet("Data");
            sheet.Cell(1, 3).Value = "Last";
            sheet.Cell(2, 3).Value = "Value";
            workbook.SaveAs(stream);
        }
        stream.Position = 0;
        using var parser = new ExcelParser(stream, null, new CsvReaderOptions { HasHeader = false });
        Assert.True(parser.Reader.Read());
        Assert.Equal(3, parser.Reader.FieldCount);
        Assert.Equal("", parser.Reader.GetString(0));
        Assert.Equal("Last", parser.Reader.GetString(2));
        Assert.True(parser.Reader.Read());
        Assert.Equal("Value", parser.Reader.GetString(2));
        Assert.False(parser.Reader.Read());
    }

    [Fact]
    public void DisposesOwnedStreamWhenWorksheetIsMissing()
    {
        var stream = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            workbook.AddWorksheet("Data");
            workbook.SaveAs(stream);
        }
        stream.Position = 0;
        Assert.Throws<ArgumentException>(() => new ExcelParser(stream, "Missing", CultureInfo.InvariantCulture));
        Assert.False(stream.CanRead);
    }

    [Fact]
    public void UsesOpenXml3ToReadAndWriteWorkbook()
    {
        using var stream = new MemoryStream();
        using (var writer = new ExcelWriter(stream, CultureInfo.InvariantCulture, leaveOpen: true))
        {
            writer.Writer.WriteRecords(new[] { new Person { Name = "OpenXML", Age = 3 } });
        }
        Assert.Equal(3, typeof(SpreadsheetDocument).Assembly.GetName().Version.Major);
        stream.Position = 0;
        using (var document = SpreadsheetDocument.Open(stream, false))
        {
            Assert.NotEmpty(document.WorkbookPart.Workbook.ChildElements);
        }
        stream.Position = 0;
        using var parser = new ExcelParser(stream, CultureInfo.InvariantCulture);
        Assert.Equal("OpenXML", Assert.Single(parser.Reader.GetRecords<Person>()).Name);
#if NET10_0_OR_GREATER
        // Loading Clippit in this same process guards the dependency combination from #1.
        Assert.Equal("Clippit", typeof(Clippit.Word.DocumentBuilder).Assembly.GetName().Name);
#endif
    }

    [Fact]
    public void FormatsNumericCellsUsingReaderCulture()
    {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.AddWorksheet("Data");
            sheet.Cell(1, 1).Value = "Amount";
            sheet.Cell(2, 1).Value = 12.5;
            workbook.SaveAs(stream);
        }
        stream.Position = 0;
        using var parser = new ExcelParser(stream, CultureInfo.GetCultureInfo("fr-FR"));
        Assert.Equal(12.5m, Assert.Single(parser.Reader.GetRecords<AmountRecord>()).Amount);
    }

    [Fact]
    public void NormalizesTransportOptionsWithoutMutatingCallerOptions()
    {
        using var stream = new MemoryStream();
        var writerOptions = new CsvWriterOptions { Separator = ';', FormatProvider = CultureInfo.GetCultureInfo("fr-FR") };
        using (var writer = new ExcelWriter(stream, "Data", writerOptions, leaveOpen: true))
        {
            writer.Writer.WriteRecords(new[] { new AmountRecord { Amount = 12.5m } });
        }
        stream.Position = 0;
        var readerOptions = new CsvReaderOptions { Separator = '|', FormatProvider = CultureInfo.GetCultureInfo("fr-FR") };
        using var parser = new ExcelParser(stream, "Data", readerOptions);
        Assert.Equal(12.5m, Assert.Single(parser.Reader.GetRecords<AmountRecord>()).Amount);
        Assert.Equal(';', writerOptions.Separator);
        Assert.Equal('|', readerOptions.Separator);
    }

    [Fact]
    public void TrimsQuotedCellWhitespaceWhenRequested()
    {
        using var stream = new MemoryStream();
        using (var writer = new ExcelWriter(stream, CultureInfo.InvariantCulture, leaveOpen: true))
        {
            writer.Writer.WriteRecords(new[] { new Person { Name = "  text, with comma  ", Age = 1 } });
        }
        stream.Position = 0;
        using var parser = new ExcelParser(stream, null, new CsvReaderOptions { TrimWhitespace = true });
        Assert.Equal("text, with comma", Assert.Single(parser.Reader.GetRecords<Person>()).Name);
    }
}

[CsvRecord]
public partial class AmountRecord
{
    public decimal Amount { get; set; }
}

